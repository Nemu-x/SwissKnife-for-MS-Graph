package actions

import (
	"crypto/rand"
	"encoding/json"
	"errors"
	"math/big"
	"net/url"
	"strings"

	"swissknife-app/internal/engine"
)

// Incident-response building blocks as catalog actions, so workflow packs can
// chain them: reset MFA, reset the password, disable suspicious inbox rules.

// AuthMethodSegments maps an authentication method's @odata.type to its
// type-specific collection (deletable methods only; the password is not).
var AuthMethodSegments = map[string]string{
	"#microsoft.graph.phoneAuthenticationMethod":                   "phoneMethods",
	"#microsoft.graph.microsoftAuthenticatorAuthenticationMethod":  "microsoftAuthenticatorMethods",
	"#microsoft.graph.softwareOathAuthenticationMethod":            "softwareOathMethods",
	"#microsoft.graph.fido2AuthenticationMethod":                   "fido2Methods",
	"#microsoft.graph.windowsHelloForBusinessAuthenticationMethod": "windowsHelloForBusinessMethods",
	"#microsoft.graph.emailAuthenticationMethod":                   "emailMethods",
	"#microsoft.graph.temporaryAccessPassAuthenticationMethod":     "temporaryAccessPassMethods",
}

// TempPassword makes a 20-character password from an alphabet without
// look-alike characters, with every character class present.
func TempPassword() (string, error) {
	const (
		lower = "abcdefghijkmnpqrstuvwxyz"
		upper = "ABCDEFGHJKLMNPQRSTUVWXYZ"
		digit = "23456789"
		sym   = "!#$%&*+-=?@"
	)
	all := lower + upper + digit + sym
	pick := func(set string) (byte, error) {
		n, err := rand.Int(rand.Reader, big.NewInt(int64(len(set))))
		if err != nil {
			return 0, err
		}
		return set[n.Int64()], nil
	}
	out := make([]byte, 0, 20)
	for _, set := range []string{lower, upper, digit, sym} {
		c, err := pick(set)
		if err != nil {
			return "", err
		}
		out = append(out, c)
	}
	for len(out) < 20 {
		c, err := pick(all)
		if err != nil {
			return "", err
		}
		out = append(out, c)
	}
	// Shuffle so the class order does not leak.
	for i := len(out) - 1; i > 0; i-- {
		j, err := rand.Int(rand.Reader, big.NewInt(int64(i+1)))
		if err != nil {
			return "", err
		}
		out[i], out[j.Int64()] = out[j.Int64()], out[i]
	}
	return string(out), nil
}

func incidentActions() []engine.Action {
	user := engine.Field{Name: "user", Kind: engine.FieldUser, Required: true}
	return []engine.Action{
		{Manifest: engine.Manifest{ID: "user.resetMfa", Page: "security", Danger: engine.Destructive, ConfirmField: "user",
			Fields: []engine.Field{user}, Permissions: []string{"UserAuthenticationMethod.ReadWrite.All"}},
			Impls: []engine.Impl{graphResetMfa{}}},
		{Manifest: engine.Manifest{ID: "user.resetPassword", Page: "security", Danger: engine.Destructive, ConfirmField: "user",
			Fields: []engine.Field{user}, Permissions: []string{"User.ReadWrite.All"}},
			Impls: []engine.Impl{graphResetPassword{}}},
		{Manifest: engine.Manifest{ID: "mail.disableRules", Page: "security", Danger: engine.Write,
			Fields: []engine.Field{user}, Permissions: []string{"MailboxSettings.ReadWrite", "Mail.ReadBasic.All", "Domain.Read.All"}},
			Impls: []engine.Impl{graphDisableRules{}}},
	}
}

// --- user.resetMfa --------------------------------------------------------

type graphResetMfa struct{}

func (graphResetMfa) Backend() engine.Backend { return engine.BackendGraph }

func (graphResetMfa) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	items, err := env.Graph.ListAll(env.Ctx, userPath(u.ID)+"/authentication/methods", nil, 0)
	if err != nil {
		return nil, err
	}
	var changes []engine.Change
	for _, raw := range items {
		var m struct {
			ID   string `json:"id"`
			Type string `json:"@odata.type"`
		}
		if json.Unmarshal(raw, &m) != nil {
			continue
		}
		seg, ok := AuthMethodSegments[m.Type]
		if !ok {
			continue // the password, or not removable
		}
		kind := strings.TrimSuffix(strings.TrimPrefix(m.Type, "#microsoft.graph."), "AuthenticationMethod")
		changes = append(changes, engine.Change{Target: u.UPN, Field: "authMethod", Op: "remove", Before: kind,
			Ref: map[string]string{"path": userPath(u.ID) + "/authentication/" + seg + "/" + url.PathEscape(m.ID)}})
	}
	if len(changes) == 0 {
		changes = []engine.Change{{Target: u.UPN, Field: "authMethod", Op: "none"}}
	}
	return changes, nil
}

func (graphResetMfa) Apply(env engine.Env, _ engine.Inputs, ch engine.Change) error {
	return env.Graph.Delete(env.Ctx, ch.Ref["path"])
}

// --- user.resetPassword ---------------------------------------------------

type graphResetPassword struct{}

func (graphResetPassword) Backend() engine.Backend { return engine.BackendGraph }

func (graphResetPassword) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	// The new password is random and never shown: after a compromise the
	// user comes back through a Temporary Access Pass or self-service reset.
	return []engine.Change{{Target: u.UPN, Field: "password", Op: "set", After: "randomPassword", Ref: map[string]string{"id": u.ID}}}, nil
}

func (graphResetPassword) Apply(env engine.Env, _ engine.Inputs, ch engine.Change) error {
	pw, err := TempPassword()
	if err != nil {
		return err
	}
	return env.Graph.Patch(env.Ctx, userPath(ch.Ref["id"]), map[string]any{
		"passwordProfile": map[string]any{"password": pw, "forceChangePasswordNextSignIn": true},
	}, nil)
}

// --- mail.disableRules ----------------------------------------------------

type graphDisableRules struct{}

func (graphDisableRules) Backend() engine.Backend { return engine.BackendGraph }

func (graphDisableRules) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	if strings.TrimSpace(in["user"]) == "" {
		return nil, errors.New("choose the mailbox")
	}
	found, err := graphRuleAudit{}.Read(env, engine.Inputs{"user": in["user"]})
	if err != nil {
		return nil, err
	}
	var changes []engine.Change
	for _, row := range found.Rows {
		if row["ruleId"] == "" {
			continue
		}
		ch := engine.Change{Target: row["rule"], Field: "inboxRule", Op: "set", Before: "enabled", After: "disabled",
			Ref: map[string]string{"path": "/users/" + url.PathEscape(row["user"]) + "/mailFolders/inbox/messageRules/" + url.PathEscape(row["ruleId"])}}
		if row["enabled"] != "yes" {
			ch.Op, ch.Before = "none", "disabled"
		}
		changes = append(changes, ch)
	}
	if len(changes) == 0 {
		changes = []engine.Change{{Target: in["user"], Field: "inboxRule", Op: "none", Note: "noSuspiciousRules"}}
	}
	return changes, nil
}

func (graphDisableRules) Apply(env engine.Env, _ engine.Inputs, ch engine.Change) error {
	return env.Graph.Patch(env.Ctx, ch.Ref["path"], map[string]any{"isEnabled": false}, nil)
}
