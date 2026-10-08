package actions

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/go-ldap/ldap/v3"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/ldapx"
)

// On-premises Active Directory over LDAP (ADR-008 M5). The inputs name AD
// objects (sAMAccountName, UPN, mail or DN), not Entra ones, so they are text
// fields: the tenant pickers would offer the wrong directory.

func adActions() []engine.Action {
	user := engine.Field{Name: "adUser", Kind: engine.FieldText, Required: true}
	return []engine.Action{
		{Manifest: engine.Manifest{ID: "ad.findUser", Page: "onprem", Danger: engine.Read,
			Fields: []engine.Field{{Name: "adQuery", Kind: engine.FieldText, Required: true}}},
			Impls: []engine.Impl{engine.ReadImpl(adFind{})}},
		{Manifest: engine.Manifest{ID: "ad.userState", Page: "onprem", Danger: engine.Write,
			Fields: []engine.Field{user, {Name: "adState", Kind: engine.FieldChoice, Required: true, Options: []string{"disable", "enable"}, Default: "disable"}}},
			Impls: []engine.Impl{adUserState{}}},
		{Manifest: engine.Manifest{ID: "ad.unlock", Page: "onprem", Danger: engine.Write, Fields: []engine.Field{user}},
			Impls: []engine.Impl{adUnlock{}}},
		{Manifest: engine.Manifest{ID: "ad.resetPassword", Page: "onprem", Danger: engine.Destructive, ConfirmField: "adUser",
			Fields: []engine.Field{user,
				{Name: "newPassword", Kind: engine.FieldSecret, Required: true},
				{Name: "mustChange", Kind: engine.FieldChoice, Required: true, Options: []string{"yes", "no"}, Default: "yes"}}},
			Impls: []engine.Impl{adResetPassword{}}},
		{Manifest: engine.Manifest{ID: "ad.groupMembership", Page: "onprem", Danger: engine.Write,
			Fields: []engine.Field{user, {Name: "adGroup", Kind: engine.FieldText, Required: true},
				{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"}}},
			Impls: []engine.Impl{adGroupMembership{}}},
	}
}

var errNoDirectory = errors.New("not connected to an on-prem directory")

func dir(env engine.Env) (*ldapx.Client, error) {
	if env.LDAP == nil {
		return nil, errNoDirectory
	}
	return env.LDAP, nil
}

// adName is how a user shows in a plan.
func adName(e ldapx.Entry) string {
	for _, k := range []string{"userPrincipalName", "sAMAccountName", "displayName"} {
		if v := e.Get(k); v != "" {
			return v
		}
	}
	return e.DN
}

// --- ad.findUser ---------------------------------------------------------

type adFind struct{}

func (adFind) Backend() engine.Backend { return engine.BackendLDAP }

func (adFind) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	c, err := dir(env)
	if err != nil {
		return nil, err
	}
	q := strings.TrimSpace(in["adQuery"])
	if q == "" {
		return nil, errors.New("type part of a name or account")
	}
	v := "*" + ldap.EscapeFilter(q) + "*"
	filter := fmt.Sprintf("(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=%s)(userPrincipalName=%s)(displayName=%s)(mail=%s)))", v, v, v, v)
	list, err := c.Search(env.Ctx, filter, ldapx.UserAttrs, 50)
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"name", "account", "upn", "enabled", "locked", "lastSignIn", "dn"}}
	for _, e := range list {
		last := ""
		if t := ldapx.FileTime(e.Get("lastLogonTimestamp")); !t.IsZero() {
			last = t.Format("2006-01-02")
		}
		res.Rows = append(res.Rows, engine.Row{
			"name": e.Get("displayName"), "account": e.Get("sAMAccountName"), "upn": e.Get("userPrincipalName"),
			"enabled": yesNo(ldapx.UAC(e)&ldapx.UACDisabled == 0), "locked": yesNo(ldapx.Locked(e)),
			"lastSignIn": last, "dn": e.DN,
		})
	}
	if len(list) == 50 {
		res.Note = &engine.Reason{Key: "firstFifty"}
	}
	return res, nil
}

// --- ad.userState --------------------------------------------------------

type adUserState struct{}

func (adUserState) Backend() engine.Backend { return engine.BackendLDAP }

func (adUserState) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	c, err := dir(env)
	if err != nil {
		return nil, err
	}
	u, err := c.FindUser(env.Ctx, in["adUser"])
	if err != nil {
		return nil, err
	}
	before := "enabled"
	if ldapx.UAC(u)&ldapx.UACDisabled != 0 {
		before = "disabled"
	}
	after := "enabled"
	if in["adState"] == "disable" {
		after = "disabled"
	}
	ch := engine.Change{Target: adName(u), Field: "adAccount", Op: "set", Before: before, After: after, Ref: map[string]string{"dn": u.DN}}
	if before == after {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (adUserState) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	c, err := dir(env)
	if err != nil {
		return err
	}
	// Re-read: other flags may have changed since the preview.
	u, err := c.FindUser(env.Ctx, ch.Ref["dn"])
	if err != nil {
		return err
	}
	uac := ldapx.UAC(u)
	if ch.After == "disabled" {
		uac |= ldapx.UACDisabled
	} else {
		uac &^= ldapx.UACDisabled
	}
	return c.Modify(env.Ctx, u.DN, ldapx.Mod{Op: "replace", Attr: "userAccountControl", Values: []string{strconv.Itoa(uac)}})
}

// --- ad.unlock -----------------------------------------------------------

type adUnlock struct{}

func (adUnlock) Backend() engine.Backend { return engine.BackendLDAP }

func (adUnlock) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	c, err := dir(env)
	if err != nil {
		return nil, err
	}
	u, err := c.FindUser(env.Ctx, in["adUser"])
	if err != nil {
		return nil, err
	}
	ch := engine.Change{Target: adName(u), Field: "adLockout", Op: "set", Before: "locked", After: "unlocked", Ref: map[string]string{"dn": u.DN}}
	if !ldapx.Locked(u) {
		ch.Op, ch.Before = "none", "unlocked"
	}
	return []engine.Change{ch}, nil
}

func (adUnlock) Apply(env engine.Env, _ engine.Inputs, ch engine.Change) error {
	c, err := dir(env)
	if err != nil {
		return err
	}
	return c.Modify(env.Ctx, ch.Ref["dn"], ldapx.Mod{Op: "replace", Attr: "lockoutTime", Values: []string{"0"}})
}

// --- ad.resetPassword ----------------------------------------------------

type adResetPassword struct{}

// Encrypted connections only: AD refuses unicodePwd over plain LDAP, and the
// catalog says so before anyone types a password.
func (adResetPassword) Backend() engine.Backend { return engine.BackendLDAPTLS }

func (adResetPassword) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	c, err := dir(env)
	if err != nil {
		return nil, err
	}
	u, err := c.FindUser(env.Ctx, in["adUser"])
	if err != nil {
		return nil, err
	}
	// The password itself never goes into the plan, journal or audit.
	after := "newPassword"
	if in["mustChange"] == "yes" {
		after = "newPasswordMustChange"
	}
	return []engine.Change{{Target: adName(u), Field: "adPassword", Op: "set", After: after, Ref: map[string]string{"dn": u.DN}}}, nil
}

func (adResetPassword) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	c, err := dir(env)
	if err != nil {
		return err
	}
	return c.SetPassword(env.Ctx, ch.Ref["dn"], in["newPassword"], in["mustChange"] == "yes")
}

// --- ad.groupMembership --------------------------------------------------

type adGroupMembership struct{}

func (adGroupMembership) Backend() engine.Backend { return engine.BackendLDAP }

func (adGroupMembership) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	c, err := dir(env)
	if err != nil {
		return nil, err
	}
	u, err := c.FindUser(env.Ctx, in["adUser"])
	if err != nil {
		return nil, err
	}
	g, err := c.FindGroup(env.Ctx, in["adGroup"])
	if err != nil {
		return nil, err
	}
	// Direct membership only, as in the group's member attribute.
	has := false
	for _, m := range g.All("member") {
		if strings.EqualFold(m, u.DN) {
			has = true
			break
		}
	}
	name := g.Get("cn")
	if name == "" {
		name = g.DN
	}
	ch := engine.Change{Target: adName(u), Field: "group.member", Op: in["op"], Ref: map[string]string{"user": u.DN, "group": g.DN}}
	if in["op"] == "add" {
		ch.After = name
		if has {
			ch.Op, ch.Before = "none", name
		}
	} else {
		ch.Before = name
		if !has {
			ch.Op, ch.Before = "none", ""
		}
	}
	return []engine.Change{ch}, nil
}

func (adGroupMembership) Apply(env engine.Env, _ engine.Inputs, ch engine.Change) error {
	c, err := dir(env)
	if err != nil {
		return err
	}
	op := "add"
	if ch.Op == "remove" {
		op = "delete"
	}
	return c.Modify(env.Ctx, ch.Ref["group"], ldapx.Mod{Op: op, Attr: "member", Values: []string{ch.Ref["user"]}})
}
