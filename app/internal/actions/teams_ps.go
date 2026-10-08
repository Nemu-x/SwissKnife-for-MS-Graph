package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/pwsh"
)

// Teams policies through MicrosoftTeams PowerShell: Graph has no API for
// meeting, messaging, calling or app policies and their assignments.

// teamsPolicy maps a field option to the policy's PowerShell names.
type teamsPolicy struct {
	userProp string // property on Get-CsOnlineUser
	typeName string // -PolicyType for group assignments
	get      string // Get-CsTeams…Policy
	grant    string // Grant-CsTeams…Policy
}

var teamsPolicies = map[string]teamsPolicy{
	"meeting":       {"TeamsMeetingPolicy", "TeamsMeetingPolicy", "Get-CsTeamsMeetingPolicy", "Grant-CsTeamsMeetingPolicy"},
	"messaging":     {"TeamsMessagingPolicy", "TeamsMessagingPolicy", "Get-CsTeamsMessagingPolicy", "Grant-CsTeamsMessagingPolicy"},
	"calling":       {"TeamsCallingPolicy", "TeamsCallingPolicy", "Get-CsTeamsCallingPolicy", "Grant-CsTeamsCallingPolicy"},
	"appSetup":      {"TeamsAppSetupPolicy", "TeamsAppSetupPolicy", "Get-CsTeamsAppSetupPolicy", "Grant-CsTeamsAppSetupPolicy"},
	"appPermission": {"TeamsAppPermissionPolicy", "TeamsAppPermissionPolicy", "Get-CsTeamsAppPermissionPolicy", "Grant-CsTeamsAppPermissionPolicy"},
}

// policyTypes is the option order the UI shows.
var policyTypes = []string{"meeting", "messaging", "calling", "appSetup", "appPermission"}

func teamsCmdlets() []string {
	out := []string{"Get-CsOnlineUser", "Get-CsGroupPolicyAssignment", "New-CsGroupPolicyAssignment",
		"Set-CsGroupPolicyAssignment", "Remove-CsGroupPolicyAssignment"}
	for _, p := range teamsPolicies {
		out = append(out, p.get, p.grant)
	}
	return out
}

func teams(env engine.Env, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error) {
	if env.PS == nil {
		return nil, errors.New("the PowerShell backend is not set up")
	}
	return env.PS.Invoke(env, pwsh.FamilyTeams, cmdlet, params, sel...)
}

// global is how the UI names "no specific policy — inherit the org default".
const global = "Global"

// policyName reads a policy reference that comes back as a plain name, as an
// object with a Name, or as null (= the Global default).
func policyName(raw json.RawMessage) string {
	var s string
	if json.Unmarshal(raw, &s) == nil && s != "" {
		return strings.TrimPrefix(s, "Tag:")
	}
	var o struct{ Name string }
	if json.Unmarshal(raw, &o) == nil && o.Name != "" {
		return strings.TrimPrefix(o.Name, "Tag:")
	}
	return global
}

// policyExists checks the name against the tenant's policies of the type.
func policyExists(env engine.Env, p teamsPolicy, name string) error {
	if strings.EqualFold(name, global) {
		return nil
	}
	rows, err := teams(env, p.get, map[string]any{"Identity": "Tag:" + name}, "Identity")
	if err != nil {
		return err
	}
	if len(rows) == 0 {
		return fmt.Errorf("no %s policy named %q", p.typeName, name)
	}
	return nil
}

type tpsBase struct{}

func (tpsBase) Backend() engine.Backend { return pwsh.BackendTeamsPS }

func teamsActions() []engine.Action {
	policyType := engine.Field{Name: "policyType", Kind: engine.FieldChoice, Required: true, Options: policyTypes, Default: "meeting"}
	policy := engine.Field{Name: "policy", Kind: engine.FieldText, Required: true, Default: global}
	perms := []string{"Teams Administrator (Entra role) for the app or the signed-in admin"}
	return []engine.Action{
		{Manifest: engine.Manifest{ID: "teams.userPolicy", Page: "teams", Danger: engine.Write,
			Fields: []engine.Field{{Name: "user", Kind: engine.FieldUser, Required: true}, policyType, policy}, Permissions: perms},
			Impls: []engine.Impl{tpsUserPolicy{}}},
		{Manifest: engine.Manifest{ID: "teams.groupPolicy", Page: "teams", Danger: engine.Write,
			Fields: []engine.Field{{Name: "group", Kind: engine.FieldGroup, Required: true}, policyType, policy}, Permissions: perms},
			Impls: []engine.Impl{tpsGroupPolicy{}}},
		{Manifest: engine.Manifest{ID: "teams.policies", Page: "teams", Danger: engine.Read,
			Fields: []engine.Field{policyType}, Permissions: perms},
			Impls: []engine.Impl{engine.ReadImpl(tpsListPolicies{})}},
		{Manifest: engine.Manifest{ID: "teams.effectivePolicies", Page: "teams", Danger: engine.Read,
			Fields: []engine.Field{{Name: "user", Kind: engine.FieldUser, Required: true}}, Permissions: perms},
			Impls: []engine.Impl{engine.ReadImpl(tpsEffectivePolicies{})}},
	}
}

// --- teams.userPolicy ---------------------------------------------------------

type tpsUserPolicy struct{ tpsBase }

func (tpsUserPolicy) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	p := teamsPolicies[in["policyType"]]
	user, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	want := strings.TrimSpace(in["policy"])
	if err := policyExists(env, p, want); err != nil {
		return nil, err
	}
	rows, err := teams(env, "Get-CsOnlineUser", map[string]any{"Identity": user.UPN}, p.userProp)
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("%s is not enabled for Teams", user.UPN)
	}
	var u map[string]json.RawMessage
	_ = json.Unmarshal(rows[0], &u)
	current := policyName(u[p.userProp])
	ch := engine.Change{Target: user.UPN, Field: "teamsPolicy." + in["policyType"], Op: "set", Before: current, After: want,
		Ref: map[string]string{"user": user.UPN}}
	if strings.EqualFold(current, want) {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (tpsUserPolicy) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	p := teamsPolicies[in["policyType"]]
	var name any = ch.After
	if strings.EqualFold(ch.After, global) {
		name = nil // back to the org-wide default
	}
	_, err := teams(env, p.grant, map[string]any{"Identity": ch.Ref["user"], "PolicyName": name})
	return err
}

// --- teams.groupPolicy ---------------------------------------------------------

type tpsGroupPolicy struct{ tpsBase }

type groupAssignment struct {
	GroupID    string `json:"GroupId"`
	PolicyName string `json:"PolicyName"`
	Rank       int    `json:"Rank"`
}

func (tpsGroupPolicy) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	p := teamsPolicies[in["policyType"]]
	var g struct {
		ID          string `json:"id"`
		DisplayName string `json:"displayName"`
	}
	if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"id,displayName"}}, &g); err != nil {
		return nil, err
	}
	want := strings.TrimSpace(in["policy"])
	if err := policyExists(env, p, want); err != nil {
		return nil, err
	}
	// All assignments of the type: a new one needs the next free rank.
	rows, err := teams(env, "Get-CsGroupPolicyAssignment", map[string]any{"PolicyType": p.typeName}, "GroupId", "PolicyName", "Rank")
	if err != nil {
		return nil, err
	}
	current, maxRank := "", 0
	for _, raw := range rows {
		var a groupAssignment
		if json.Unmarshal(raw, &a) != nil {
			continue
		}
		if a.Rank > maxRank {
			maxRank = a.Rank
		}
		if strings.EqualFold(a.GroupID, g.ID) {
			current = strings.TrimPrefix(a.PolicyName, "Tag:")
		}
	}
	ref := map[string]string{"group": g.ID, "rank": fmt.Sprint(maxRank + 1)}
	target := firstOf(g.DisplayName, g.ID)
	field := "teamsPolicy." + in["policyType"]
	switch {
	case strings.EqualFold(want, global) && current == "":
		return []engine.Change{{Target: target, Field: field, Op: "none", Before: global, After: global, Ref: ref}}, nil
	case strings.EqualFold(want, global):
		return []engine.Change{{Target: target, Field: field, Op: "remove", Before: current, Ref: ref}}, nil
	case current == "":
		return []engine.Change{{Target: target, Field: field, Op: "add", After: want, Ref: ref}}, nil
	case strings.EqualFold(current, want):
		return []engine.Change{{Target: target, Field: field, Op: "none", Before: current, After: want, Ref: ref}}, nil
	default:
		return []engine.Change{{Target: target, Field: field, Op: "set", Before: current, After: want, Ref: ref}}, nil
	}
}

func (tpsGroupPolicy) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	p := teamsPolicies[in["policyType"]]
	params := map[string]any{"GroupId": ch.Ref["group"], "PolicyType": p.typeName}
	cmdlet := "Remove-CsGroupPolicyAssignment"
	switch ch.Op {
	case "add":
		cmdlet, params["PolicyName"], params["Rank"] = "New-CsGroupPolicyAssignment", ch.After, ch.Ref["rank"]
	case "set":
		cmdlet, params["PolicyName"] = "Set-CsGroupPolicyAssignment", ch.After
	}
	_, err := teams(env, cmdlet, params)
	return err
}

// --- reads -------------------------------------------------------------------

type tpsListPolicies struct{ tpsBase }

func (tpsListPolicies) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	p := teamsPolicies[in["policyType"]]
	rows, err := teams(env, p.get, nil, "Identity", "Description")
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"policy", "description"}}
	for _, raw := range rows {
		var r struct{ Identity, Description string }
		if json.Unmarshal(raw, &r) == nil {
			res.Rows = append(res.Rows, engine.Row{"policy": strings.TrimPrefix(r.Identity, "Tag:"), "description": r.Description})
		}
	}
	return res, nil
}

type tpsEffectivePolicies struct{ tpsBase }

func (tpsEffectivePolicies) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	user, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	sel := make([]string, 0, len(policyTypes))
	for _, t := range policyTypes {
		sel = append(sel, teamsPolicies[t].userProp)
	}
	rows, err := teams(env, "Get-CsOnlineUser", map[string]any{"Identity": user.UPN}, sel...)
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("%s is not enabled for Teams", user.UPN)
	}
	var u map[string]json.RawMessage
	_ = json.Unmarshal(rows[0], &u)
	res := &engine.ReadResult{Columns: []string{"policyType", "policy"}}
	for _, t := range policyTypes {
		res.Rows = append(res.Rows, engine.Row{"policyType": t, "policy": policyName(u[teamsPolicies[t].userProp])})
	}
	return res, nil
}
