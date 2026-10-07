// Package actions holds the built-in catalog actions (ADR-008). Each action
// pairs a manifest with implementations; this file has the Microsoft Graph
// ones. Values in Change.Before/After are display values: object names, or
// tokens the frontend translates under actions.values.*.
package actions

import (
	"encoding/json"
	"errors"
	"net/url"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// Builtin returns every built-in action, for Engine.Register.
func Builtin() []engine.Action {
	return []engine.Action{
		{
			Manifest: engine.Manifest{
				ID: "user.signIn", Page: "users", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "state", Kind: engine.FieldChoice, Required: true, Options: []string{"blocked", "allowed"}, Default: "blocked"},
				},
				Permissions: []string{"User.ReadWrite.All"},
			},
			Impls: []engine.Impl{graphSignIn{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "user.revokeSessions", Page: "users", Danger: engine.Destructive,
				Fields:       []engine.Field{{Name: "user", Kind: engine.FieldUser, Required: true}},
				ConfirmField: "user",
				Permissions:  []string{"User.ReadWrite.All"},
			},
			Impls: []engine.Impl{graphRevokeSessions{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "user.manager", Page: "users", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "manager", Kind: engine.FieldUser, Required: true},
				},
				Permissions: []string{"User.ReadWrite.All"},
			},
			Impls: []engine.Impl{graphManager{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "user.usageLocation", Page: "users", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "country", Kind: engine.FieldText, Required: true},
				},
				Permissions: []string{"User.ReadWrite.All"},
			},
			Impls: []engine.Impl{graphUsageLocation{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "group.membership", Page: "groups", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "group", Kind: engine.FieldGroup, Required: true},
					{Name: "role", Kind: engine.FieldChoice, Required: true, Options: []string{"member", "owner"}, Default: "member"},
					{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
				},
				Permissions: []string{"GroupMember.ReadWrite.All", "User.Read.All"},
			},
			Impls: []engine.Impl{graphGroupMembership{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "license.assign", Page: "licensing", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "sku", Kind: engine.FieldSku, Required: true},
					{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
				},
				Permissions: []string{"User.ReadWrite.All"},
			},
			Impls: []engine.Impl{graphLicense{}},
		},
	}
}

type graphUser struct {
	ID               string `json:"id"`
	UPN              string `json:"userPrincipalName"`
	UsageLocation    string `json:"usageLocation"`
	AccountEnabled   *bool  `json:"accountEnabled"`
	LicenseStates    []struct {
		SkuID           string  `json:"skuId"`
		AssignedByGroup *string `json:"assignedByGroup"`
	} `json:"licenseAssignmentStates"`
}

func getUser(env engine.Env, user, sel string) (*graphUser, error) {
	var u graphUser
	err := env.Graph.Get(env.Ctx, "/users/"+url.PathEscape(user), url.Values{"$select": {sel}}, &u)
	if err != nil {
		return nil, err
	}
	if u.UPN == "" {
		u.UPN = user
	}
	return &u, nil
}

func userPath(id string) string { return "/users/" + url.PathEscape(id) }

// --- user.signIn ---------------------------------------------------------

type graphSignIn struct{}

func (graphSignIn) Backend() engine.Backend { return engine.BackendGraph }

func (graphSignIn) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName,accountEnabled")
	if err != nil {
		return nil, err
	}
	before := "allowed"
	if u.AccountEnabled != nil && !*u.AccountEnabled {
		before = "blocked"
	}
	ch := engine.Change{Target: u.UPN, Field: "signIn", Op: "set", Before: before, After: in["state"],
		Ref: map[string]string{"id": u.ID}}
	if before == in["state"] {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (graphSignIn) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	return env.Graph.Patch(env.Ctx, userPath(ch.Ref["id"]), map[string]any{"accountEnabled": ch.After == "allowed"}, nil)
}

// --- user.revokeSessions -------------------------------------------------

type graphRevokeSessions struct{}

func (graphRevokeSessions) Backend() engine.Backend { return engine.BackendGraph }

func (graphRevokeSessions) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	// Graph cannot tell whether sessions exist; revoking is always a change.
	return []engine.Change{{Target: u.UPN, Field: "sessions", Op: "set", Before: "active", After: "revoked",
		Ref: map[string]string{"id": u.ID}}}, nil
}

func (graphRevokeSessions) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	return env.Graph.Post(env.Ctx, userPath(ch.Ref["id"])+"/revokeSignInSessions", map[string]any{}, nil)
}

// --- group.membership ----------------------------------------------------

type graphGroupMembership struct{}

func (graphGroupMembership) Backend() engine.Backend { return engine.BackendGraph }

func (graphGroupMembership) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	var g struct {
		ID          string `json:"id"`
		DisplayName string `json:"displayName"`
	}
	if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"id,displayName"}}, &g); err != nil {
		return nil, err
	}
	// Direct membership / ownership only: memberOf lists direct groups, and
	// ownedObjects what the user owns — a transitive member is still added.
	nav := "/memberOf"
	if in["role"] == "owner" {
		nav = "/ownedObjects"
	}
	items, err := env.Graph.ListAll(env.Ctx, userPath(u.ID)+nav, url.Values{"$select": {"id"}}, 0)
	if err != nil {
		return nil, err
	}
	has := false
	for _, raw := range items {
		var o struct {
			ID string `json:"id"`
		}
		if json.Unmarshal(raw, &o) == nil && o.ID == g.ID {
			has = true
			break
		}
	}
	ch := engine.Change{Target: u.UPN, Field: "group." + in["role"], Op: in["op"],
		Ref: map[string]string{"user": u.ID, "group": g.ID}}
	name := g.DisplayName
	if name == "" {
		name = g.ID
	}
	if in["op"] == "add" {
		ch.After = name
		if has {
			ch.Op, ch.Before = "none", name
		}
	} else {
		ch.Before = name
		if !has {
			ch.Op, ch.Before, ch.After = "none", "", ""
		}
	}
	return []engine.Change{ch}, nil
}

func (graphGroupMembership) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	kind := "members"
	if in["role"] == "owner" {
		kind = "owners"
	}
	base := "/groups/" + url.PathEscape(ch.Ref["group"]) + "/" + kind
	if ch.Op == "remove" {
		return env.Graph.Delete(env.Ctx, base+"/"+url.PathEscape(ch.Ref["user"])+"/$ref")
	}
	body := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + ch.Ref["user"]}
	return env.Graph.Post(env.Ctx, base+"/$ref", body, nil)
}

// --- license.assign ------------------------------------------------------

type graphLicense struct{}

func (graphLicense) Backend() engine.Backend { return engine.BackendGraph }

func (graphLicense) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName,licenseAssignmentStates")
	if err != nil {
		return nil, err
	}
	// Name the SKU the way the license pages do (part number); the frontend
	// turns it into the friendly product name.
	name := in["sku"]
	skus, err := env.Graph.ListAll(env.Ctx, "/subscribedSkus", url.Values{"$select": {"skuId,skuPartNumber"}}, 0)
	if err != nil {
		return nil, err
	}
	for _, raw := range skus {
		var s struct {
			SkuID string `json:"skuId"`
			Part  string `json:"skuPartNumber"`
		}
		if json.Unmarshal(raw, &s) == nil && strings.EqualFold(s.SkuID, in["sku"]) {
			name = s.Part
		}
	}
	// A license can be held directly and/or through groups; only a direct
	// assignment can be removed here (Graph rejects the rest).
	direct, inherited := false, false
	for _, l := range u.LicenseStates {
		if !strings.EqualFold(l.SkuID, in["sku"]) {
			continue
		}
		if l.AssignedByGroup == nil || *l.AssignedByGroup == "" {
			direct = true
		} else {
			inherited = true
		}
	}
	ch := engine.Change{Target: u.UPN, Field: "license", Op: in["op"], Ref: map[string]string{"id": u.ID}}
	switch {
	case in["op"] == "add" && (direct || inherited):
		ch.Op, ch.Before, ch.After = "none", name, name
	case in["op"] == "add":
		ch.After = name
	case direct:
		ch.Before = name
	case inherited:
		ch.Op, ch.Before, ch.Note = "none", name, "inheritedLicense"
	default:
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (graphLicense) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	add, remove := []map[string]any{}, []string{}
	if ch.Op == "add" {
		add = append(add, map[string]any{"skuId": in["sku"]})
	} else {
		remove = append(remove, in["sku"])
	}
	return env.Graph.Post(env.Ctx, userPath(ch.Ref["id"])+"/assignLicense",
		map[string]any{"addLicenses": add, "removeLicenses": remove}, nil)
}

// --- user.manager --------------------------------------------------------

type graphManager struct{}

func (graphManager) Backend() engine.Backend { return engine.BackendGraph }

func (graphManager) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	m, err := getUser(env, in["manager"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	if m.ID == u.ID {
		return nil, errors.New("a user cannot be their own manager")
	}
	// No manager answers 404; anything else is a real failure.
	// The manager may be an organizational contact without a UPN.
	var cur struct {
		ID   string `json:"id"`
		UPN  string `json:"userPrincipalName"`
		Name string `json:"displayName"`
	}
	err = env.Graph.Get(env.Ctx, userPath(u.ID)+"/manager", url.Values{"$select": {"id,userPrincipalName,displayName"}}, &cur)
	if err != nil && !isNotFound(err) {
		return nil, err
	}
	ch := engine.Change{Target: u.UPN, Field: "manager", Op: "set", Before: firstOf(cur.UPN, cur.Name, cur.ID), After: m.UPN,
		Ref: map[string]string{"id": u.ID, "manager": m.ID}}
	if cur.ID == m.ID {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (graphManager) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	body := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/users/" + ch.Ref["manager"]}
	return env.Graph.Put(env.Ctx, userPath(ch.Ref["id"])+"/manager/$ref", body, nil)
}

// --- user.usageLocation --------------------------------------------------

type graphUsageLocation struct{}

func (graphUsageLocation) Backend() engine.Backend { return engine.BackendGraph }

func (graphUsageLocation) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	loc := strings.ToUpper(strings.TrimSpace(in["country"]))
	if len(loc) != 2 || strings.Trim(loc, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") != "" {
		return nil, errors.New("usage location must be a two-letter country code, e.g. US or DE")
	}
	u, err := getUser(env, in["user"], "id,userPrincipalName,usageLocation")
	if err != nil {
		return nil, err
	}
	ch := engine.Change{Target: u.UPN, Field: "usageLocation", Op: "set", Before: u.UsageLocation, After: loc,
		Ref: map[string]string{"id": u.ID}}
	if strings.EqualFold(u.UsageLocation, loc) {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (graphUsageLocation) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	return env.Graph.Patch(env.Ctx, userPath(ch.Ref["id"]), map[string]any{"usageLocation": ch.After}, nil)
}

func isNotFound(err error) bool {
	var ge *graphapi.GraphError
	return errors.As(err, &ge) && ge.StatusCode == 404
}

func firstOf(vals ...string) string {
	for _, v := range vals {
		if v != "" {
			return v
		}
	}
	return ""
}
