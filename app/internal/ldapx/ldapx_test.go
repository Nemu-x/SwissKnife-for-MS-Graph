package ldapx_test

import (
	"context"
	"errors"
	"strings"
	"testing"

	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/ldapx/ldaptest"
)

const (
	base   = "dc=corp,dc=example"
	bindDN = "cn=svc,ou=service,dc=corp,dc=example"
	annDN  = "cn=Ann Lee,ou=staff,dc=corp,dc=example"
)

func directory(t *testing.T, ldaps bool) *ldaptest.Server {
	return ldaptest.Start(t, bindDN, "s3cret", ldaps, map[string]map[string][]string{
		annDN: {"objectClass": {"user"}, "sAMAccountName": {"ann"}, "userPrincipalName": {"ann@corp.example"},
			"displayName": {"Ann Lee"}, "userAccountControl": {"512"}, "lockoutTime": {"133000000000000000"}},
		"cn=Sales,ou=groups,dc=corp,dc=example": {"objectClass": {"group"}, "cn": {"Sales"}, "sAMAccountName": {"sales"}},
	})
}

func client(s *ldaptest.Server, mode ldapx.TLSMode, password string) *ldapx.Client {
	c := ldapx.New(ldapx.Config{Host: s.Host, Port: s.Port, TLS: mode, BaseDN: base, BindDN: bindDN}, password)
	c.TLSConfig = s.Client
	return c
}

func TestFindModifyAndResetOverLDAPS(t *testing.T) {
	s := directory(t, true)
	c := client(s, ldapx.LDAPS, "s3cret")
	ctx := context.Background()
	if err := client(s, ldapx.LDAPS, "wrong").Test(ctx); err == nil {
		t.Fatal("a wrong bind password must fail")
	}
	u, err := c.FindUser(ctx, "ann@corp.example")
	if err != nil || u.DN != annDN || !ldapx.Locked(u) || ldapx.UAC(u)&ldapx.UACDisabled != 0 {
		t.Fatalf("find: %+v %v", u, err)
	}
	if _, err := c.FindUser(ctx, "nobody"); err == nil || !strings.Contains(err.Error(), "not found") {
		t.Fatalf("missing user: %v", err)
	}
	if err := c.Modify(ctx, annDN, ldapx.Mod{Op: "replace", Attr: "lockoutTime", Values: []string{"0"}}); err != nil {
		t.Fatal(err)
	}
	if got := s.Attr(annDN, "lockoutTime"); len(got) != 1 || got[0] != "0" {
		t.Fatalf("lockoutTime %q", got)
	}
	if err := c.SetPassword(ctx, annDN, "N3w-Passw0rd!", true); err != nil {
		t.Fatal(err)
	}
	if got := s.Attr(annDN, "unicodePwd"); len(got) != 1 || got[0] != ldapx.UnicodePwd("N3w-Passw0rd!") {
		t.Fatalf("unicodePwd not set as quoted UTF-16LE")
	}
	if got := s.Attr(annDN, "pwdLastSet"); len(got) != 1 || got[0] != "0" {
		t.Fatalf("must-change not set: %v", got)
	}
	g, err := c.FindGroup(ctx, "Sales")
	if err != nil || !strings.HasPrefix(g.DN, "cn=Sales") {
		t.Fatalf("group %+v %v", g, err)
	}
}

func TestStartTLSAndPlainRefusesPasswordReset(t *testing.T) {
	s := directory(t, false)
	ctx := context.Background()
	if _, err := client(s, ldapx.StartTLS, "s3cret").FindUser(ctx, "ann"); err != nil {
		t.Fatalf("StartTLS: %v", err)
	}
	plain := client(s, ldapx.Plain, "s3cret")
	if _, err := plain.FindUser(ctx, "ann"); err != nil {
		t.Fatalf("plain: %v", err)
	}
	if err := plain.SetPassword(ctx, annDN, "x", false); !errors.Is(err, ldapx.ErrNeedsTLS) {
		t.Fatalf("plain LDAP must refuse password resets: %v", err)
	}
}

func TestUnicodePwdAndFileTime(t *testing.T) {
	if got := []byte(ldapx.UnicodePwd("ab")); string(got) != "\"\x00a\x00b\x00\"\x00" {
		t.Fatalf("encoding %q", got)
	}
	if ldapx.FileTime("0").IsZero() != true || ldapx.FileTime("116444736000000000").Unix() != 0 {
		t.Fatal("file time conversion")
	}
}
