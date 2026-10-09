package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"reflect"
	"sort"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

func (x snapshotRestore) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	doc, err := x.svc.load(in["snapshot"])
	if err != nil {
		return nil, err
	}
	if err := sameTenant(doc.Meta, env.TenantID); err != nil {
		return nil, fmt.Errorf("%w; restore only into the tenant it came from", err)
	}
	filter := strings.TrimSpace(in["object"])
	var changes []engine.Change
	if sec := in["section"]; sec != "all" {
		if changes, err = x.planSection(env, doc, sec, filter, nil); err != nil {
			return nil, err
		}
		if len(changes) == 0 && filter != "" {
			return nil, fmt.Errorf("no object named %q in the snapshot's %s", filter, sec)
		}
	} else {
		if filter != "" {
			return nil, errors.New("restoring every section takes no object name")
		}
		// Every section the snapshot holds, one step each. A section that
		// cannot be read now (no PowerShell, no permission) is reported and
		// left out; the others still restore.
		recreated := map[string]bool{} // snapshot named-location ids this plan recreates
		step := 0
		for _, sec := range restoreOrder {
			if _, ok := doc.Sections[sec]; !ok {
				continue
			}
			step++
			chs, err := x.planSection(env, doc, sec, "", recreated)
			if err != nil {
				chs = []engine.Change{{Target: sec, Field: "restoredSection", Op: "none", Note: "sectionSkipped", After: err.Error()}}
			}
			for i := range chs {
				chs[i].Step, chs[i].StepNo = "snapshot.section."+sec, step
				if sec == "namedLocations" && chs[i].Op == "add" {
					recreated[chs[i].Ref["id"]] = true
				}
			}
			changes = append(changes, chs...)
		}
		if step == 0 {
			return nil, errors.New("the snapshot holds no section that can be restored")
		}
	}
	// The typed confirmation names how many objects will be rewritten.
	n := 0
	for _, c := range changes {
		if c.Op != "none" {
			n++
		}
	}
	// On a real change: the engine ignores it on rows that change nothing.
	for i := range changes {
		if changes[i].Op != "none" {
			changes[i].Ref[engine.ConfirmRef] = fmt.Sprintf("restore %d", n)
			break
		}
	}
	return changes, nil
}

// encryptedSettings reports a custom profile holding encrypted OMA-URI
// values: Graph reads them back without the value, so writing the profile
// would overwrite the secret.
func encryptedSettings(o map[string]any) bool {
	list, _ := o["omaSettings"].([]any)
	for _, s := range list {
		if m, ok := s.(map[string]any); ok {
			if enc, _ := m["isEncrypted"].(bool); enc {
				return true
			}
		}
	}
	return false
}

// liveGroups drops assignments to groups deleted since the snapshot: Graph
// refuses the whole assign call for one missing group. It reports whether
// any was dropped.
func liveGroups(env engine.Env, targets []any, seen map[string]bool) ([]any, bool, error) {
	out := targets[:0:0]
	dropped := false
	for _, t := range targets {
		g, _ := t.(map[string]any)["target"].(map[string]any)["groupId"].(string)
		if g != "" {
			ok, known := seen[g]
			if !known {
				err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(g), url.Values{"$select": {"id"}}, nil)
				var ge *graphapi.GraphError
				switch {
				case err == nil:
					ok = true
				case errors.As(err, &ge) && ge.StatusCode == 404:
				default:
					return nil, false, err
				}
				seen[g] = ok
			}
			if !ok {
				dropped = true
				continue
			}
		}
		out = append(out, t)
	}
	return out, dropped, nil
}

// planSection plans one section. recreated, in an "all" restore, names the
// snapshot's named locations recreated earlier in the same plan: policies
// that point at them are remapped when applied.
func (x snapshotRestore) planSection(env engine.Env, doc *snapshotFile, sec, filter string, recreated map[string]bool) ([]engine.Change, error) {
	objs, ok := doc.Sections[sec]
	if !ok {
		return nil, fmt.Errorf("the snapshot has no %s section (it was skipped when taken)", sec)
	}
	if pr, ok := psRestorable[sec]; ok {
		return x.planPS(env, sec, pr, objs, filter)
	}
	rs, ok := restorableSections[sec]
	if !ok {
		return nil, fmt.Errorf("section %s cannot be restored", sec)
	}
	byName, err := liveByName(env, rs.path)
	if err != nil {
		return nil, err
	}
	// Policies refer to named locations by id; resolve them against today.
	var snapLoc map[string]string
	var liveLoc map[string][]string
	liveLocIDs := map[string]bool{}
	if sec == "conditionalAccessPolicies" {
		snapLoc = map[string]string{}
		for _, l := range doc.Sections["namedLocations"] {
			id, _ := l["id"].(string)
			name, _ := l["displayName"].(string)
			snapLoc[id] = name
		}
		if liveLoc, err = liveByName(env, "/identity/conditionalAccess/namedLocations"); err != nil {
			return nil, err
		}
		for _, ids := range liveLoc {
			for _, id := range ids {
				liveLocIDs[id] = true
			}
		}
	}
	var params url.Values
	if rs.assign {
		params = url.Values{"$expand": {"assignments"}}
	}
	var changes []engine.Change
	groups := map[string]bool{}
	for _, o := range objs {
		id, _ := o["id"].(string)
		label := labelOf(o)
		if filter != "" && !strings.EqualFold(label, filter) && id != filter {
			continue
		}
		if encryptedSettings(o) {
			changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "none", Note: "encryptedSettings",
				Ref: map[string]string{"section": sec, "id": id}})
			continue
		}
		norm := normalizeValue(o).(map[string]any)
		want := rs.subset(norm)
		ref := map[string]string{"section": sec, "id": id, "path": rs.path}
		if snapLoc != nil {
			if missing := remapLocations(want, snapLoc, liveLoc, liveLocIDs); len(missing) > 0 {
				for _, m := range missing {
					if !recreated[m] {
						return nil, fmt.Errorf("policy %q uses named locations that no longer exist (%s); restore the named locations first", label, strings.Join(missing, ", "))
					}
				}
				// Recreated in this plan: their new ids are known only then.
				names, _ := json.Marshal(snapLoc)
				ref["remap"] = string(names)
			}
		}
		wantAssign := assignmentTargets(norm["assignments"])
		note := ""
		if rs.assign {
			var dropped bool
			if wantAssign, dropped, err = liveGroups(env, wantAssign, groups); err != nil {
				return nil, err
			}
			if dropped {
				note = "assignmentGroupGone"
			}
		}
		if rs.assign && len(wantAssign) > 0 {
			a, _ := json.Marshal(wantAssign)
			ref["assign"] = string(a)
		}

		var live map[string]any
		err := env.Graph.Get(env.Ctx, rs.path+"/"+url.PathEscape(id), params, &live)
		var ge *graphapi.GraphError
		switch {
		case errors.As(err, &ge) && ge.StatusCode == 404:
			// Recreated by hand under the same name: rewrite that one rather
			// than adding a duplicate.
			name, _ := o["displayName"].(string)
			others := byName[strings.ToLower(name)]
			if len(others) > 1 {
				return nil, fmt.Errorf("%q was deleted and several objects now have that name; restore it by hand", name)
			}
			if len(others) == 1 {
				other := others[0]
				ref["id"] = other
				if err := env.Graph.Get(env.Ctx, rs.path+"/"+url.PathEscape(other), params, &live); err != nil {
					return nil, err
				}
				break
			}
			for k, v := range rs.create {
				want[k] = v
			}
			if acts := scheduledActions(o["scheduledActionsForRule"]); acts != nil && rs.create != nil {
				want["scheduledActionsForRule"] = acts
			} else if rs.create != nil && note == "" {
				note = "defaultComplianceAction"
			}
			body, _ := json.Marshal(want)
			ref["body"] = string(body)
			changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "add", After: "snapshot", Note: note, Ref: ref})
			continue
		case err != nil:
			return nil, err
		}
		body, _ := json.Marshal(want)
		ref["body"] = string(body)
		liveNorm := normalizeValue(live).(map[string]any)
		have := rs.subset(liveNorm)
		keys := rs.writable
		if keys == nil {
			for k := range want {
				if k != "@odata.type" {
					keys = append(keys, k)
				}
			}
		}
		var differ []string
		for _, k := range keys {
			if !reflect.DeepEqual(normalizeValue(want[k]), have[k]) {
				differ = append(differ, k)
			}
		}
		if len(differ) > 0 {
			ref["patch"] = "1"
		}
		if rs.assign {
			if reflect.DeepEqual(wantAssign, assignmentTargets(liveNorm["assignments"])) {
				delete(ref, "assign")
			} else {
				differ = append(differ, "assignments")
				if ref["assign"] == "" {
					ref["assign"] = "[]" // assigned since the snapshot: unassign
				}
			}
		}
		if len(differ) == 0 {
			if filter != "" {
				changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "none", Ref: ref})
			}
			continue
		}
		sort.Strings(differ)
		changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "set",
			Before: strings.Join(differ, ", "), After: "snapshot", Note: note, Ref: ref})
	}
	return changes, nil
}

// scheduledActions turns a snapshot's non-compliance actions into what a
// create takes (ids dropped); nil when the snapshot has none.
func scheduledActions(v any) []any {
	list, _ := v.([]any)
	if len(list) == 0 {
		return nil
	}
	var out []any
	for _, r := range list {
		rule, _ := r.(map[string]any)
		if rule == nil {
			continue
		}
		var confs []any
		cl, _ := rule["scheduledActionConfigurations"].([]any)
		for _, c := range cl {
			if m, ok := c.(map[string]any); ok {
				cp := map[string]any{}
				for k, x := range m {
					if k != "id" && !strings.HasPrefix(k, "@") {
						cp[k] = x
					}
				}
				confs = append(confs, cp)
			}
		}
		out = append(out, map[string]any{"ruleName": rule["ruleName"], "scheduledActionConfigurations": confs})
	}
	return out
}

// assignmentTargets reduces Intune assignments to their targets, in a stable
// order (assignment ids change with every assign call).
func assignmentTargets(v any) []any {
	list, _ := v.([]any)
	out := []any{}
	for _, a := range list {
		if m, ok := a.(map[string]any); ok && m["target"] != nil {
			out = append(out, map[string]any{"target": normalizeValue(m["target"])})
		}
	}
	sort.Slice(out, func(i, j int) bool {
		a, _ := json.Marshal(out[i])
		b, _ := json.Marshal(out[j])
		return string(a) < string(b)
	})
	return out
}

// psLabel names a PowerShell object the way its admin center does.
func psLabel(o map[string]any) string {
	for _, k := range []string{"Name", "Identity"} {
		if s, ok := o[k].(string); ok && s != "" {
			return s
		}
	}
	return labelOf(o)
}

func psSectionOf(name string) *psSection {
	for _, sc := range snapshotSections {
		if sc.name == name {
			return sc.ps
		}
	}
	return nil
}

// planPS compares a PowerShell section with what its cmdlet reads now.
func (x snapshotRestore) planPS(env engine.Env, sec string, pr psRestore, objs []map[string]any, filter string) ([]engine.Change, error) {
	live, err := x.svc.collectPS(env.Ctx, psSectionOf(sec))
	if err != nil {
		return nil, err
	}
	liveByID := map[string]map[string]any{}
	for _, o := range live {
		id, _ := o["id"].(string)
		if pr.global {
			id = ""
		}
		liveByID[strings.ToLower(id)] = o
	}
	var changes []engine.Change
	for _, o := range objs {
		id, _ := o["id"].(string)
		label := psLabel(o)
		if filter != "" && !strings.EqualFold(label, filter) && !strings.EqualFold(id, filter) {
			continue
		}
		key := strings.ToLower(id)
		if pr.global {
			key = ""
		}
		params := map[string]any{}
		ref := map[string]string{"section": sec, "id": id}
		cur, ok := liveByID[key]
		if !ok {
			if pr.create == "" {
				// The snapshot keeps too little to rebuild it (a transport
				// rule's conditions): say so rather than pretend.
				changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "none", Note: "cannotRecreate", Ref: ref})
				continue
			}
			for _, k := range pr.props {
				if v := o[k]; v != nil {
					params[k] = v
				}
			}
			b, _ := json.Marshal(params)
			ref["params"], ref["create"] = string(b), "1"
			changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "add", After: "snapshot", Ref: ref})
			continue
		}
		var differ []string
		for _, k := range pr.props {
			// A property the snapshot holds as empty is left alone: an empty
			// value cannot be told from one the cmdlet did not return.
			if v := o[k]; v != nil && !reflect.DeepEqual(normalizeValue(v), normalizeValue(cur[k])) {
				params[k] = v
				differ = append(differ, k)
			}
		}
		if pr.state && o["State"] != nil && !reflect.DeepEqual(normalizeValue(o["State"]), normalizeValue(cur["State"])) {
			ref["state"] = fmt.Sprint(o["State"])
			differ = append(differ, "State")
		}
		if len(differ) == 0 {
			if filter != "" {
				changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "none", Ref: ref})
			}
			continue
		}
		sort.Strings(differ)
		b, _ := json.Marshal(params)
		ref["params"] = string(b)
		changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "set",
			Before: strings.Join(differ, ", "), After: "snapshot", Ref: ref})
	}
	return changes, nil
}

func (x snapshotRestore) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	if pr, ok := psRestorable[ch.Ref["section"]]; ok {
		return applyPS(env, ch, pr)
	}
	var body map[string]any
	if err := json.Unmarshal([]byte(ch.Ref["body"]), &body); err != nil {
		return err
	}
	if names := ch.Ref["remap"]; names != "" {
		// Named locations recreated earlier in this run have new ids.
		var snapLoc map[string]string
		if err := json.Unmarshal([]byte(names), &snapLoc); err != nil {
			return err
		}
		live, err := liveByName(env, "/identity/conditionalAccess/namedLocations")
		if err != nil {
			return err
		}
		liveIDs := map[string]bool{}
		for _, ids := range live {
			for _, id := range ids {
				liveIDs[id] = true
			}
		}
		if missing := remapLocations(body, snapLoc, live, liveIDs); len(missing) > 0 {
			return fmt.Errorf("named locations %s were not recreated; restore them first", strings.Join(missing, ", "))
		}
	}
	id := ch.Ref["id"]
	rs := restorableSections[ch.Ref["section"]]
	switch {
	case ch.Op == "add":
		// Recreated objects get a new id; the old one is gone for good.
		var made struct {
			ID string `json:"id"`
		}
		if err := env.Graph.Post(env.Ctx, ch.Ref["path"], body, &made); err != nil {
			return err
		}
		id = made.ID
	case ch.Ref["patch"] != "" || !rs.assign:
		if err := env.Graph.Patch(env.Ctx, ch.Ref["path"]+"/"+url.PathEscape(id), body, nil); err != nil {
			return err
		}
	}
	if a := ch.Ref["assign"]; a != "" {
		if id == "" {
			return errors.New("created, but Graph returned no id: assign it by hand")
		}
		var list []any
		if err := json.Unmarshal([]byte(a), &list); err != nil {
			return err
		}
		return env.Graph.Post(env.Ctx, ch.Ref["path"]+"/"+url.PathEscape(id)+"/assign", map[string]any{"assignments": list}, nil)
	}
	return nil
}

func applyPS(env engine.Env, ch engine.Change, pr psRestore) error {
	if env.PS == nil {
		return errors.New("the PowerShell backend is not set up")
	}
	family := psSectionOf(ch.Ref["section"]).family
	params := map[string]any{}
	if p := ch.Ref["params"]; p != "" {
		if err := json.Unmarshal([]byte(p), &params); err != nil {
			return err
		}
	}
	id := ch.Ref["id"]
	if ch.Ref["create"] != "" {
		// Teams names custom policies "Tag:<name>"; New- takes the name.
		params["Identity"] = strings.TrimPrefix(id, "Tag:")
		_, err := env.PS.Invoke(env, family, pr.create, params)
		return err
	}
	if len(params) > 0 {
		if !pr.global {
			params["Identity"] = id
		}
		if _, err := env.PS.Invoke(env, family, pr.set, params); err != nil {
			return err
		}
	}
	if st := ch.Ref["state"]; st != "" {
		cmdlet := "Disable-TransportRule"
		if strings.EqualFold(st, "Enabled") {
			cmdlet = "Enable-TransportRule"
		}
		_, err := env.PS.Invoke(env, family, cmdlet, map[string]any{"Identity": id, "Confirm": false})
		return err
	}
	return nil
}
