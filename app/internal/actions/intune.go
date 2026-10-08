package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"sort"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// The Intune explorer: what lands on a person or a group, which policies are
// assigned to nobody, where Intune itself reports conflicts, and assigning or
// unassigning a policy for a group.

// policyFamily is one Intune collection that carries assignments.
type policyFamily struct {
	kind   string // display token
	path   string // collection path (v1.0 unless beta)
	beta   bool
	name   string // property holding the display name
	filter string // optional $filter to skip noise
}

var policyFamilies = []policyFamily{
	{"configuration", "/deviceManagement/deviceConfigurations", false, "displayName", ""},
	{"compliance", "/deviceManagement/deviceCompliancePolicies", false, "displayName", ""},
	{"settingsCatalog", "/deviceManagement/configurationPolicies", true, "name", ""},
	// Built-in and store apps run into thousands; only assigned ones matter.
	{"app", "/deviceAppManagement/mobileApps", false, "displayName", "isAssigned eq true"},
}

// Assignment target types, compared in full.
const (
	targetGroup      = "#microsoft.graph.groupAssignmentTarget"
	targetExclusion  = "#microsoft.graph.exclusionGroupAssignmentTarget"
	targetAllUsers   = "#microsoft.graph.allLicensedUsersAssignmentTarget"
	targetAllDevices = "#microsoft.graph.allDevicesAssignmentTarget"
)

// intuneAssignment keeps the target as it came: rebuilding the list for
// /assign must carry assignment filters and anything else unchanged.
type intuneAssignment struct {
	Target json.RawMessage `json:"target"`
	Intent string          `json:"intent,omitempty"` // apps only
}

type targetView struct {
	Type    string `json:"@odata.type"`
	GroupID string `json:"groupId"`
}

func (a intuneAssignment) view() targetView {
	var v targetView
	_ = json.Unmarshal(a.Target, &v)
	return v
}

type intunePolicy struct {
	family      policyFamily
	ID          string
	Name        string
	Modified    string
	Assignments []intuneAssignment
}

// listPolicies reads the families. A family the tenant does not have (beta
// 400/404) is skipped quietly; one the app may not read (403) is skipped and
// named in the returned list so reports can say they are incomplete; other
// errors (throttling, outages) fail the read.
func listPolicies(env engine.Env, fams []policyFamily) ([]intunePolicy, []string, error) {
	var out []intunePolicy
	var denied []string
	for _, f := range fams {
		path := f.path
		if f.beta {
			path = env.Graph.Beta(f.path)
		}
		q := url.Values{"$expand": {"assignments"}}
		if f.filter != "" {
			q.Set("$filter", f.filter)
		}
		raw, err := env.Graph.ListAll(env.Ctx, path, q, 0)
		if err != nil {
			var ge *graphapi.GraphError
			switch {
			case errors.As(err, &ge) && ge.StatusCode == 403:
				denied = append(denied, f.kind)
				continue
			case errors.As(err, &ge) && (ge.StatusCode == 400 || ge.StatusCode == 404):
				continue
			}
			return nil, nil, err
		}
		for _, r := range raw {
			var p map[string]json.RawMessage
			if json.Unmarshal(r, &p) != nil {
				continue
			}
			var id, name, modified string
			_ = json.Unmarshal(p["id"], &id)
			_ = json.Unmarshal(p[f.name], &name)
			_ = json.Unmarshal(p["lastModifiedDateTime"], &modified)
			var as []intuneAssignment
			_ = json.Unmarshal(p["assignments"], &as)
			if f.kind == "app" && len(as) == 0 {
				continue
			}
			out = append(out, intunePolicy{family: f, ID: id, Name: name, Modified: modified, Assignments: as})
		}
	}
	return out, denied, nil
}

// deniedNote names the families a report could not read.
func deniedNote(denied []string) *engine.Reason {
	if len(denied) == 0 {
		return nil
	}
	return &engine.Reason{Key: "intuneDenied", Params: map[string]string{"kinds": strings.Join(denied, ", ")}}
}

func intuneActions() []engine.Action {
	perms := []string{"DeviceManagementConfiguration.Read.All", "DeviceManagementApps.Read.All", "Group.Read.All"}
	read := func(id string, fields []engine.Field, r engine.Reader) engine.Action {
		return engine.Action{Manifest: engine.Manifest{ID: id, Page: "intune", Danger: engine.Read, Fields: fields, Permissions: perms},
			Impls: []engine.Impl{engine.ReadImpl(r)}}
	}
	return []engine.Action{
		read("intune.assignedTo", []engine.Field{{Name: "user", Kind: engine.FieldUser}, {Name: "group", Kind: engine.FieldGroup}}, graphAssignedTo{}),
		read("intune.unassigned", nil, graphUnassigned{}),
		read("intune.conflicts", nil, graphConflicts{}),
		{Manifest: engine.Manifest{ID: "intune.assign", Page: "intune", Danger: engine.Write,
			Fields: []engine.Field{
				{Name: "policyKind", Kind: engine.FieldChoice, Required: true, Options: []string{"configuration", "compliance"}, Default: "configuration"},
				{Name: "policy", Kind: engine.FieldText, Required: true},
				{Name: "group", Kind: engine.FieldGroup, Required: true},
				{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
			},
			Permissions: []string{"DeviceManagementConfiguration.ReadWrite.All"}},
			Impls: []engine.Impl{graphIntuneAssign{}}},
	}
}

// --- what lands on a user / group -----------------------------------------------------

type graphAssignedTo struct{}

func (graphAssignedTo) Backend() engine.Backend { return engine.BackendGraph }

// Read lists assignments reaching the user through their groups (device
// groups of their devices are not followed) or reaching the group.
func (graphAssignedTo) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	if (in["user"] == "") == (in["group"] == "") {
		return nil, errors.New("pick either a user or a group")
	}
	groups := map[string]string{}
	forUser := in["user"] != ""
	if forUser {
		raw, err := env.Graph.ListAll(env.Ctx, "/users/"+url.PathEscape(in["user"])+"/transitiveMemberOf/microsoft.graph.group",
			url.Values{"$select": {"id,displayName"}}, 0)
		if err != nil {
			return nil, err
		}
		for _, r := range raw {
			var g struct{ ID, DisplayName string }
			if json.Unmarshal(r, &g) == nil {
				groups[g.ID] = g.DisplayName
			}
		}
	} else {
		var g struct{ ID, DisplayName string }
		if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"id,displayName"}}, &g); err != nil {
			return nil, err
		}
		groups[g.ID] = g.DisplayName
	}
	policies, denied, err := listPolicies(env, policyFamilies)
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"kind", "policy", "via", "effect"}, Note: deniedNote(denied)}
	if res.Note == nil && forUser {
		res.Note = &engine.Reason{Key: "intuneUserGroupsOnly"}
	}
	for _, p := range policies {
		for _, a := range p.Assignments {
			t := a.view()
			via, effect := "", "included"
			switch t.Type {
			case targetExclusion:
				if name, ok := groups[t.GroupID]; ok {
					via, effect = name, "excluded"
				}
			case targetGroup:
				if name, ok := groups[t.GroupID]; ok {
					via = name
				}
			case targetAllUsers:
				if forUser {
					via = "allUsers"
				}
			case targetAllDevices:
				via = "allDevices"
			}
			if via == "" {
				continue
			}
			if a.Intent != "" && effect == "included" {
				effect = a.Intent // apps: required / available / uninstall
			}
			res.Rows = append(res.Rows, engine.Row{"kind": p.family.kind, "policy": p.Name, "via": via, "effect": effect})
		}
	}
	sort.SliceStable(res.Rows, func(i, j int) bool { return res.Rows[i]["kind"] < res.Rows[j]["kind"] })
	return res, nil
}

// --- assigned to nobody ------------------------------------------------------------------

type graphUnassigned struct{}

func (graphUnassigned) Backend() engine.Backend { return engine.BackendGraph }

func (graphUnassigned) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	policies, denied, err := listPolicies(env, policyFamilies[:3]) // apps are listed only when assigned
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"kind", "policy", "lastModified"}, Note: deniedNote(denied)}
	for _, p := range policies {
		if len(p.Assignments) == 0 {
			res.Rows = append(res.Rows, engine.Row{"kind": p.family.kind, "policy": p.Name, "lastModified": p.Modified})
		}
	}
	return res, nil
}

// --- conflicts as Intune reports them ---------------------------------------------------------

type graphConflicts struct{}

func (graphConflicts) Backend() engine.Backend { return engine.BackendGraph }

// Read lists device conflicts Intune reports for configuration and compliance
// policies (settings catalog and endpoint security have no such report in
// Graph). Comparing settings by hand would only guess.
func (graphConflicts) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	policies, denied, err := listPolicies(env, policyFamilies[:2])
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"kind", "policy", "device", "user", "status"}}
	for _, p := range policies {
		if len(p.Assignments) == 0 {
			continue
		}
		raw, err := env.Graph.ListAll(env.Ctx, p.family.path+"/"+url.PathEscape(p.ID)+"/deviceStatuses",
			url.Values{"$filter": {"status eq 'conflict'"}}, 0)
		if err != nil {
			var ge *graphapi.GraphError
			if errors.As(err, &ge) && ge.StatusCode == 403 {
				denied = append(denied, p.family.kind+" device status")
				continue
			}
			return nil, err
		}
		for _, r := range raw {
			var s struct {
				DeviceDisplayName string `json:"deviceDisplayName"`
				UserPrincipalName string `json:"userPrincipalName"`
				Status            string `json:"status"`
			}
			// The $filter is not honoured everywhere: check the status too.
			if json.Unmarshal(r, &s) == nil && s.Status == "conflict" {
				res.Rows = append(res.Rows, engine.Row{"kind": p.family.kind, "policy": p.Name,
					"device": s.DeviceDisplayName, "user": s.UserPrincipalName, "status": s.Status})
			}
		}
	}
	res.Note = deniedNote(denied)
	return res, nil
}

// --- assign / unassign a policy for a group --------------------------------------------------

type graphIntuneAssign struct{}

func (graphIntuneAssign) Backend() engine.Backend { return engine.BackendGraph }

func familyOf(kind string) (policyFamily, error) {
	for _, f := range policyFamilies[:2] {
		if f.kind == kind {
			return f, nil
		}
	}
	return policyFamily{}, fmt.Errorf("policies of type %q cannot be assigned here", kind)
}

// findPolicy resolves the policy by its exact (case-insensitive) name.
func findPolicy(env engine.Env, in engine.Inputs) (*intunePolicy, error) {
	fam, err := familyOf(in["policyKind"])
	if err != nil {
		return nil, err
	}
	policies, denied, err := listPolicies(env, []policyFamily{fam})
	if err != nil {
		return nil, err
	}
	if len(denied) > 0 {
		return nil, fmt.Errorf("no permission to read %s policies", fam.kind)
	}
	var hit []intunePolicy
	for _, p := range policies {
		if strings.EqualFold(p.Name, strings.TrimSpace(in["policy"])) {
			hit = append(hit, p)
		}
	}
	switch len(hit) {
	case 0:
		return nil, fmt.Errorf("no %s policy named %q", fam.kind, in["policy"])
	case 1:
		return &hit[0], nil
	default:
		return nil, fmt.Errorf("%d %s policies are named %q — rename one first", len(hit), fam.kind, in["policy"])
	}
}

// rebuild returns the assignment list to send: every existing target exactly
// as it is (filters included), with the group's include target added or
// removed. /assign replaces the whole list, so nothing may be lost here.
func rebuild(p *intunePolicy, groupID, op string) (next []map[string]json.RawMessage, has bool, err error) {
	for _, a := range p.Assignments {
		t := a.view()
		if t.GroupID == groupID {
			switch t.Type {
			case targetGroup:
				has = true
				if op == "remove" {
					continue
				}
			case targetExclusion:
				if op == "add" {
					return nil, false, errors.New("this group is excluded from the policy — remove the exclusion in Intune first")
				}
			}
		}
		next = append(next, map[string]json.RawMessage{"target": a.Target})
	}
	if op == "add" && !has {
		t, _ := json.Marshal(map[string]string{"@odata.type": targetGroup, "groupId": groupID})
		next = append(next, map[string]json.RawMessage{"target": t})
	}
	return next, has, nil
}

func (graphIntuneAssign) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	p, err := findPolicy(env, in)
	if err != nil {
		return nil, err
	}
	var g struct{ ID, DisplayName string }
	if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"id,displayName"}}, &g); err != nil {
		return nil, err
	}
	_, has, err := rebuild(p, g.ID, in["op"])
	if err != nil {
		return nil, err
	}
	ch := presence(p.Name, "intuneAssignment", in["op"], firstOf(g.DisplayName, g.ID), has,
		map[string]string{"policy": p.ID, "group": g.ID})
	return []engine.Change{ch}, nil
}

// Apply re-reads the policy: assignments changed in the portal since the
// preview must not be overwritten by a stale list.
func (graphIntuneAssign) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	p, err := findPolicy(env, in)
	if err != nil {
		return err
	}
	if p.ID != ch.Ref["policy"] {
		return errors.New("the policy changed since the preview — preview again")
	}
	next, _, err := rebuild(p, ch.Ref["group"], ch.Op)
	if err != nil {
		return err
	}
	return env.Graph.Post(env.Ctx, p.family.path+"/"+url.PathEscape(p.ID)+"/assign", map[string]any{"assignments": next}, nil)
}
