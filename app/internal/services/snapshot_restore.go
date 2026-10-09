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
	"swissknife-app/internal/session"
)

// Restoring configuration from a snapshot: the plan is the difference
// between the snapshot and the tenant now, per object, and applying it puts
// the snapshot's values back (or recreates what was deleted).

// restorable describes a snapshot section that can be written back.
type restorable struct {
	path     string   // collection path
	writable []string // properties sent on PATCH/POST
	typed    bool     // @odata.type must accompany writes (named locations)
}

var restorableSections = map[string]restorable{
	"conditionalAccessPolicies": {"/identity/conditionalAccess/policies",
		[]string{"displayName", "state", "conditions", "grantControls", "sessionControls"}, false},
	"namedLocations": {"/identity/conditionalAccess/namedLocations",
		[]string{"displayName", "ipRanges", "isTrusted", "countriesAndRegions", "includeUnknownCountriesAndRegions", "countryLookupMethod"}, true},
}

func snapshotRestoreAction(s *session.Session) engine.Action {
	return engine.Action{
		Manifest: engine.Manifest{
			ID: "config.restore", Capability: "config.snapshot.restore", Page: "security", Danger: engine.Destructive,
			Fields: []engine.Field{
				{Name: "snapshot", Kind: engine.FieldSnapshot, Required: true},
				{Name: "section", Kind: engine.FieldChoice, Required: true, Options: []string{"conditionalAccessPolicies", "namedLocations"}, Default: "conditionalAccessPolicies"},
				{Name: "object", Kind: engine.FieldText},
			},
			ConfirmField: "snapshot",
			Permissions:  []string{"Policy.ReadWrite.ConditionalAccess", "Policy.Read.All"},
		},
		Impls: []engine.Impl{snapshotRestore{svc: NewSnapshotService(s)}},
	}
}

type snapshotRestore struct{ svc *SnapshotService }

func (snapshotRestore) Backend() engine.Backend { return engine.BackendGraph }

// subset keeps the writable properties of an object.
func (r restorable) subset(o map[string]any) map[string]any {
	out := map[string]any{}
	for _, k := range r.writable {
		if v, ok := o[k]; ok {
			out[k] = v
		}
	}
	// An authentication strength is read back in full but written by id only.
	if g, ok := out["grantControls"].(map[string]any); ok {
		if as, ok := g["authenticationStrength"].(map[string]any); ok {
			g2 := make(map[string]any, len(g))
			for k, v := range g {
				g2[k] = v
			}
			g2["authenticationStrength"] = map[string]any{"id": as["id"]}
			out["grantControls"] = g2
		}
	}
	if r.typed {
		if t, ok := o["@odata.type"]; ok {
			out["@odata.type"] = t
		}
	}
	return out
}

// sameTenant reports whether a snapshot was taken in the connected tenant.
// Snapshots from before tenant ids were recorded never match: profile names
// are local labels and prove nothing about the directory.
func sameTenant(m SnapshotMeta, tenantID string) error {
	if m.TenantID == "" {
		return errors.New("this snapshot does not record its tenant (taken by an older version); take a new snapshot")
	}
	if tenantID == "" || !strings.EqualFold(m.TenantID, tenantID) {
		return fmt.Errorf("the snapshot was taken in another tenant (%s)", m.Tenant)
	}
	return nil
}

// liveByName indexes a collection's current objects by lower-cased display
// name; a name can belong to several objects.
func liveByName(env engine.Env, path string) (map[string][]string, error) {
	objs, err := graphapi.ListAllInto[map[string]any](env.Ctx, env.Graph, path, url.Values{"$select": {"id,displayName"}}, 0)
	if err != nil {
		return nil, err
	}
	out := map[string][]string{}
	for _, o := range objs {
		name, _ := o["displayName"].(string)
		id, _ := o["id"].(string)
		if name != "" && id != "" {
			out[strings.ToLower(name)] = append(out[strings.ToLower(name)], id)
		}
	}
	return out, nil
}

// remapLocations points a policy's location conditions at today's named
// locations: a location recreated after deletion has a new id, so the
// snapshot's id is matched through its name (only when exactly one location
// has that name). Unresolvable ids are returned.
func remapLocations(policy map[string]any, snapNames map[string]string, live map[string][]string, liveIDs map[string]bool) []string {
	cond, _ := policy["conditions"].(map[string]any)
	locs, _ := cond["locations"].(map[string]any)
	if locs == nil {
		return nil
	}
	var missing []string
	for _, k := range []string{"includeLocations", "excludeLocations"} {
		list, _ := locs[k].([]any)
		for i, v := range list {
			id, _ := v.(string)
			if id == "" || id == "All" || id == "AllTrusted" || liveIDs[id] {
				continue
			}
			if ids := live[strings.ToLower(snapNames[id])]; len(ids) == 1 && snapNames[id] != "" {
				list[i] = ids[0]
				continue
			}
			missing = append(missing, id)
		}
	}
	return missing
}

func (x snapshotRestore) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	doc, err := x.svc.load(in["snapshot"])
	if err != nil {
		return nil, err
	}
	if err := sameTenant(doc.Meta, env.TenantID); err != nil {
		return nil, fmt.Errorf("%w; restore only into the tenant it came from", err)
	}
	sec := in["section"]
	rs, ok := restorableSections[sec]
	if !ok {
		return nil, fmt.Errorf("section %s cannot be restored", sec)
	}
	objs, ok := doc.Sections[sec]
	if !ok {
		return nil, fmt.Errorf("the snapshot has no %s section (it was skipped when taken)", sec)
	}
	filter := strings.TrimSpace(in["object"])
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
	var changes []engine.Change
	for _, o := range objs {
		id, _ := o["id"].(string)
		label := labelOf(o)
		if filter != "" && !strings.EqualFold(label, filter) && id != filter {
			continue
		}
		want := rs.subset(normalizeValue(o).(map[string]any))
		if snapLoc != nil {
			if missing := remapLocations(want, snapLoc, liveLoc, liveLocIDs); len(missing) > 0 {
				return nil, fmt.Errorf("policy %q uses named locations that no longer exist (%s); restore the named locations first", label, strings.Join(missing, ", "))
			}
		}
		body, _ := json.Marshal(want)
		ref := map[string]string{"id": id, "body": string(body), "path": rs.path}

		var live map[string]any
		err := env.Graph.Get(env.Ctx, rs.path+"/"+url.PathEscape(id), nil, &live)
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
				if err := env.Graph.Get(env.Ctx, rs.path+"/"+url.PathEscape(other), nil, &live); err != nil {
					return nil, err
				}
				break
			}
			changes = append(changes, engine.Change{Target: label, Field: "restoredObject", Op: "add", After: "snapshot", Ref: ref})
			continue
		case err != nil:
			return nil, err
		}
		have := rs.subset(normalizeValue(live).(map[string]any))
		var differ []string
		for _, k := range rs.writable {
			if !reflect.DeepEqual(normalizeValue(want[k]), have[k]) {
				differ = append(differ, k)
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
			Before: strings.Join(differ, ", "), After: "snapshot", Ref: ref})
	}
	if len(changes) == 0 && filter != "" {
		return nil, fmt.Errorf("no object named %q in the snapshot's %s", filter, sec)
	}
	// The typed confirmation names how many objects will be rewritten.
	n := 0
	for _, c := range changes {
		if c.Op != "none" {
			n++
		}
	}
	if len(changes) > 0 {
		changes[0].Ref[engine.ConfirmRef] = fmt.Sprintf("restore %d", n)
	}
	return changes, nil
}

func (snapshotRestore) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	var body map[string]any
	if err := json.Unmarshal([]byte(ch.Ref["body"]), &body); err != nil {
		return err
	}
	if ch.Op == "add" {
		// Recreated objects get a new id; the old one is gone for good.
		return env.Graph.Post(env.Ctx, ch.Ref["path"], body, nil)
	}
	return env.Graph.Patch(env.Ctx, ch.Ref["path"]+"/"+url.PathEscape(ch.Ref["id"]), body, nil)
}
