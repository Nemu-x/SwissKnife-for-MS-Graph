package services

import (
	"strings"
	"testing"

	"github.com/zalando/go-keyring"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/ldapx/ldaptest"
	"swissknife-app/internal/session"
)

const (
	adBase  = "dc=corp,dc=example"
	adBind  = "cn=svc,ou=service,dc=corp,dc=example"
	adAnn   = "cn=Ann Lee,ou=staff,dc=corp,dc=example"
	adSales = "cn=Sales,ou=groups,dc=corp,dc=example"
)

func adDirectory(t *testing.T, ldaps bool) *ldaptest.Server {
	return ldaptest.Start(t, adBind, "s3cret", ldaps, map[string]map[string][]string{
		adAnn: {"objectClass": {"user"}, "sAMAccountName": {"ann"}, "userPrincipalName": {"ann@corp.example"},
			"displayName": {"Ann Lee"}, "userAccountControl": {"512"}, "lockoutTime": {"133000000000000000"}},
		adSales: {"objectClass": {"group"}, "cn": {"Sales"}},
	})
}

// No tenant is connected: the AD actions run on the directory alone.
func TestADActionsRunWithoutATenant(t *testing.T) {
	dirSrv := adDirectory(t, true)
	sess := session.New(auditlog.New(t.TempDir()))
	c := ldapx.New(ldapx.Config{ID: "c1", Host: dirSrv.Host, Port: dirSrv.Port, TLS: ldapx.LDAPS, BaseDN: adBase, BindDN: adBind}, "s3cret")
	c.TLSConfig = dirSrv.Client
	directories.Store(sess, c)
	t.Cleanup(func() { directories.Delete(sess) })
	e := NewEngine(sess)

	for _, entry := range e.Catalog() {
		if strings.HasPrefix(entry.ID, "ad.") && !entry.Available {
			t.Fatalf("%s unavailable: %+v", entry.ID, entry.Reason)
		}
	}
	rows, err := e.Run(t.Context(), "ad.findUser", engine.Inputs{"adQuery": "ann"})
	if err != nil || len(rows.Rows) != 1 || rows.Rows[0]["locked"] != "yes" {
		t.Fatalf("find: %+v %v", rows, err)
	}

	apply := func(id string, in engine.Inputs, confirm string) *engine.Plan {
		t.Helper()
		p, err := e.Plan(id, in)
		if err != nil {
			t.Fatalf("%s plan: %v", id, err)
		}
		if confirm == "" {
			confirm = p.ConfirmTarget
		}
		if _, err := e.Apply(p.ID, confirm); err != nil {
			t.Fatalf("%s apply: %v", id, err)
		}
		return p
	}
	apply("ad.userState", engine.Inputs{"adUser": "ann", "adState": "disable"}, "")
	if got := dirSrv.Attr(adAnn, "userAccountControl"); got[0] != "514" {
		t.Fatalf("disable: uac %v", got)
	}
	apply("ad.unlock", engine.Inputs{"adUser": "ann@corp.example"}, "")
	if got := dirSrv.Attr(adAnn, "lockoutTime"); got[0] != "0" {
		t.Fatalf("unlock: %v", got)
	}
	p := apply("ad.resetPassword", engine.Inputs{"adUser": "ann", "newPassword": "N3w-Passw0rd!", "mustChange": "yes"}, "ann")
	if p.Inputs["newPassword"] == "N3w-Passw0rd!" || strings.Contains(p.Changes[0].After, "N3w") {
		t.Fatal("the password must not come back in the plan")
	}
	if got := dirSrv.Attr(adAnn, "unicodePwd"); len(got) != 1 || got[0] != ldapx.UnicodePwd("N3w-Passw0rd!") {
		t.Fatal("password not set")
	}
	apply("ad.groupMembership", engine.Inputs{"adUser": "ann", "adGroup": "Sales", "op": "add"}, "")
	if got := dirSrv.Attr(adSales, "member"); len(got) != 1 || got[0] != adAnn {
		t.Fatalf("member %v", got)
	}
	p, err = e.Plan("ad.groupMembership", engine.Inputs{"adUser": "ann", "adGroup": "Sales", "op": "add"})
	if err != nil || p.Changes[0].Op != "none" {
		t.Fatalf("already a member must be a no-op: %+v %v", p, err)
	}

	// A disconnect from the directory invalidates its plans.
	p, _ = e.Plan("ad.unlock", engine.Inputs{"adUser": "ann"})
	directories.Delete(sess)
	if _, err := e.Apply(p.ID, ""); err == nil {
		t.Fatal("a plan must not outlive its directory connection")
	}
}

func TestPlainLDAPHidesPasswordReset(t *testing.T) {
	dirSrv := adDirectory(t, false)
	sess := session.New(auditlog.New(t.TempDir()))
	directories.Store(sess, ldapx.New(ldapx.Config{Host: dirSrv.Host, Port: dirSrv.Port, TLS: ldapx.Plain, BaseDN: adBase, BindDN: adBind}, "s3cret"))
	t.Cleanup(func() { directories.Delete(sess) })
	for _, entry := range NewEngine(sess).Catalog() {
		if entry.ID == "ad.resetPassword" && (entry.Available || entry.Reason == nil || entry.Reason.Key != "ldapNeedsTLS") {
			t.Fatalf("reset over plain LDAP: %+v", entry)
		}
		if entry.ID == "ad.unlock" && !entry.Available {
			t.Fatalf("unlock works over plain LDAP: %+v", entry)
		}
	}
}

func TestOnPremConnectionsKeepThePasswordInTheKeychain(t *testing.T) {
	keyring.MockInit()
	dirSrv := adDirectory(t, false)
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(t.TempDir())
	o := NewOnPremService(sess)
	if _, err := o.SaveConnection(ldapx.Config{Host: dirSrv.Host, TLS: "tls13"}, ""); err == nil {
		t.Fatal("an unknown TLS mode must be refused")
	}
	c, err := o.SaveConnection(ldapx.Config{Name: "Lab", Host: dirSrv.Host, Port: dirSrv.Port, TLS: ldapx.Plain, BaseDN: adBase, BindDN: adBind}, "s3cret")
	if err != nil || !c.HasPassword {
		t.Fatalf("save: %+v %v", c, err)
	}
	if st, err := o.Connect(c.ID); err != nil || !st.Connected || st.Secure {
		t.Fatalf("connect: %+v %v", st, err)
	}
	if directoryFor(sess) == nil {
		t.Fatal("connect must make the directory active")
	}
	if err := o.DeleteConnection(c.ID); err != nil {
		t.Fatal(err)
	}
	if directoryFor(sess) != nil {
		t.Fatal("deleting the active connection disconnects it")
	}
	if _, ok := keyringGet(onpremKey(c.ID)); ok {
		t.Fatal("the password must be removed with the connection")
	}
	list, _ := o.Connections()
	if len(list) != 0 {
		t.Fatalf("connections %+v", list)
	}
}

func keyringGet(k string) (string, bool) {
	v, err := keyring.Get("SwissKnifeGraph", k)
	return v, err == nil
}

// A tenant profile's limits (here: read only, scoped) do not govern the
// on-prem directory; the read-only switch does.
func TestTenantLimitsDoNotGovernTheDirectory(t *testing.T) {
	dirSrv := ldaptest.Start(t, adBind, "s3cret", true, map[string]map[string][]string{
		adAnn: {"objectClass": {"user"}, "sAMAccountName": {"ann"}, "userAccountControl": {"66048"}, "lockoutTime": {"5"}, "primaryGroupID": {"513"}},
		"cn=Domain Users,ou=groups,dc=corp,dc=example": {"objectClass": {"group"}, "cn": {"Domain Users"}, "primaryGroupToken": {"513"}},
	})
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetPolicy("p1", session.Policy{MaxDanger: "read", AllowedGroups: []string{"g1"}})
	c := ldapx.New(ldapx.Config{Name: "CORP", Host: dirSrv.Host, Port: dirSrv.Port, TLS: ldapx.LDAPS, BaseDN: adBase, BindDN: adBind}, "s3cret")
	c.TLSConfig = dirSrv.Client
	directories.Store(sess, c)
	t.Cleanup(func() { directories.Delete(sess) })
	e := NewEngine(sess)

	p, err := e.Plan("ad.unlock", engine.Inputs{"adUser": "ann"})
	if err != nil {
		t.Fatalf("tenant limits must not block AD: %v", err)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatalf("apply: %v", err)
	}
	if _, err := e.Plan("ad.groupMembership", engine.Inputs{"adUser": "ann", "adGroup": "Domain Users", "op": "remove"}); err == nil || !strings.Contains(err.Error(), "primary group") {
		t.Fatalf("primary group: %v", err)
	}
	// 66048 = normal account + password never expires.
	if _, err := e.Plan("ad.resetPassword", engine.Inputs{"adUser": "ann", "newPassword": "x", "mustChange": "yes"}); err == nil || !strings.Contains(err.Error(), "never expires") {
		t.Fatalf("must-change on a never-expiring password: %v", err)
	}
	sess.SetReadOnly(true)
	p, _ = e.Plan("ad.userState", engine.Inputs{"adUser": "ann", "adState": "disable"})
	if _, err := e.Apply(p.ID, ""); err == nil {
		t.Fatal("the read-only switch covers the directory")
	}
}
