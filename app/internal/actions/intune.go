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
	assign bool   // supports the bulk assign action
}

var policyFamilies = []policyFamily{
	{"configuration", "/deviceManagement/deviceConfigurations", false, "displayName", true},
	{"compliance", "/deviceManagement/deviceCompliancePolicies", false, "displayName", true},
	{"settingsCatalog", "/deviceManagement/configurationPolicies", true, "name", false},
	{"app", "/deviceAppManagement/mobileApps", false, "displayName", false},
}

type intuneAssignment struct {
	ID     string `json:"id,omitempty"`
	Target struct {
		Type    string `json:"@odata.type"`
		GroupID string `json:"groupId,omitempty"`
	} `json:"target"`
	Intent string `json:"intent,omitempty"` // apps only
}

type intunePolicy struct {
	family      policyFamily
	ID          string
	Name        string
	Modified    string
	Assignments []intuneAssignment
}

func listPolicies(env engine.Env, fams []policyFamily) ([]intunePolicy, error) {
	var out []intunePolicy
	for _, f := range fams {
		path := f.path
		if f.beta {
			path = env.Graph.Beta(f.path)
		}
		raw, err := env.Graph.ListAll(env.Ctx, path, url.Values{"$expand": {"assignments"}}, 0)
		if err != nil {
			// Settings catalog is beta-only and a tenant without Intune
			// licenses answers 403/400 for some families: skip, not fail.
			var ge *graphapi.GraphError
			if errors.As(err, &ge) && (ge.StatusCode == 400 || ge.StatusCode == 403 || ge.StatusCode == 404) {
				continue
			}
			return nil, err
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
			// Apps without any assignment are just catalogue entries.
			if f.kind == "app" && len(as) == 0 {
				continue
			}
			out = append(out, intunePolicy{family: f, ID: id, Name: name, Modified: modified, Assignments: as})
		}
	}
	return out, nil
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

func (graphAssignedTo) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	if (in["user"] == "") == (in["group"] == "") {
		return nil, errors.New("pick either a user or a group")
	}
	// The groups whose assignments reach the target, with names for display.
	groups := map[string]string{}
	allUsers := in["user"] != ""
	if in["user"] != "" {
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
	policies, err := listPolicies(env, policyFamilies)
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"kind", "policy", "via", "effect"}}
	for _, p := range policies {
		for _, a := range p.Assignments {
			via, effect := "", "included"
			switch {
			case strings.HasSuffix(a.Target.Type, "exclusionGroupAssignmentTarget"):
				if name, ok := groups[a.Target.GroupID]; ok {
					via, effect = name, "excluded"
				}
			case strings.HasSuffix(a.Target.Type, "groupAssignmentTarget"):
				if name, ok := groups[a.Target.GroupID]; ok {
					via = name
				}
			case strings.HasSuffix(a.Target.Type, "allLicensedUsersAssignmentTarget") && allUsers:
				via = "allUsers"
			case strings.HasSuffix(a.Target.Type, "allDevicesAssignmentTarget"):
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
	policies, err := listPolicies(env, policyFamilies[:3]) // apps without assignments are skipped anyway
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"kind", "policy", "lastModified"}}
	for _, p := range policies {
		if len(p.Assignments) == 0 {
			res.Rows = append(res.Rows, engine.Row{"kind": p.family.kind, "policy": p.Name, "lastModified": firstOf(p.Modified, "")})
		}
	}
	return res, nil
}

// --- conflicts as Intune reports them ---------------------------------------------------------

type graphConflicts struct{}

func (graphConflicts) Backend() engine.Backend { return engine.BackendGraph }

// Comparing thousands of settings by hand would guess; Intune already marks
// devices whose policies disagree. This lists those reports.
func (graphConflicts) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	policies, err := listPolicies(env, policyFamilies[:2])
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
			continue // a policy type without device statuses
		}
		for _, r := range raw {
			var s struct {
				DeviceDisplayName string `json:"deviceDisplayName"`
				UserPrincipalName string `json:"userPrincipalName"`
				Status            string `json:"status"`
			}
			if json.Unmarshal(r, &s) == nil {
				res.Rows = append(res.Rows, engine.Row{"kind": p.family.kind, "policy": p.Name,
					"device": s.DeviceDisplayName, "user": s.UserPrincipalName, "status": s.Status})
			}
		}
	}
	return res, nil
}

// --- assign / unassign a policy for a group --------------------------------------------------

type graphIntuneAssign struct{}

func (graphIntuneAssign) Backend() engine.Backend { return engine.BackendGraph }

func familyOf(kind string) policyFamily {
	for _, f := range policyFamilies {
		if f.kind == kind {
			return f
		}
	}
	return policyFamilies[0]
}

func (graphIntuneAssign) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	fam := familyOf(in["policyKind"])
	policies, err := listPolicies(env, []policyFamily{fam})
	if err != nil {
		return nil, err
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
	default:
		return nil, fmt.Errorf("%d %s policies are named %q — rename one first", len(hit), fam.kind, in["policy"])
	}
	p := hit[0]
	var g struct{ ID, DisplayName string }
	if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"id,displayName"}}, &g); err != nil {
		return nil, err
	}
	has := false
	for _, a := range p.Assignments {
		if a.Target.GroupID == g.ID && strings.HasSuffix(a.Target.Type, ".groupAssignmentTarget") {
			has = true
		}
	}
	// assign replaces the whole list: the plan carries the list to send.
	next := make([]map[string]any, 0, len(p.Assignments)+1)
	for _, a := range p.Assignments {
		if in["op"] == "remove" && a.Target.GroupID == g.ID && strings.HasSuffix(a.Target.Type, ".groupAssignmentTarget") {
			continue
		}
		// Everything else stays as it was: all-users/all-devices targets carry
		// no group id, exclusions keep theirs.
		target := map[string]any{"@odata.type": a.Target.Type}
		if a.Target.GroupID != "" {
			target["groupId"] = a.Target.GroupID
		}
		next = append(next, map[string]any{"target": target})
	}
	if in["op"] == "add" && !has {
		next = append(next, map[string]any{"target": map[string]any{"@odata.type": "#microsoft.graph.groupAssignmentTarget", "groupId": g.ID}})
	}
	body, _ := json.Marshal(next)
	ch := presence(p.Name, "intuneAssignment", in["op"], firstOf(g.DisplayName, g.ID), has,
		map[string]string{"path": fam.path + "/" + url.PathEscape(p.ID) + "/assign", "assignments": string(body)})
	return []engine.Change{ch}, nil
}

func (graphIntuneAssign) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	var list []map[string]any
	if err := json.Unmarshal([]byte(ch.Ref["assignments"]), &list); err != nil {
		return err
	}
	return env.Graph.Post(env.Ctx, ch.Ref["path"], map[string]any{"assignments": list}, nil)
}
