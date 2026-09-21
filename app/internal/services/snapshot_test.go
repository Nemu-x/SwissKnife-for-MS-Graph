package services

import (
	"encoding/json"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// snapState is the mutable "tenant" behind the fake: tests change it between
// two snapshots to produce drift.
type snapState struct {
	caPolicies string
	members    string
}

func defaultSnapState() *snapState {
	return &snapState{
		caPolicies: `{"value":[
			{"@odata.context":"ctx","id":"p2","displayName":"Legacy auth","state":"enabled","modifiedDateTime":"2026-01-01T00:00:00Z"},
			{"id":"p1","displayName":"MFA for all","state":"enabled","modifiedDateTime":"2026-01-01T00:00:00Z",
			 "conditions":{"users":{"includeUsers":["All"],"excludeUsers":["u9","u8"]}},
			 "grantControls":{"operator":"OR","builtInControls":["mfa"]}}]}`,
		members: `{"value":[{"id":"u1","userPrincipalName":"admin@contoso.com","displayName":"Admin"}]}`,
	}
}

func snapshotHarness(t *testing.T, st *snapState, calls *[]string) *SnapshotService {
	t.Helper()
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if calls != nil {
			q := r.URL.RawQuery
			if q != "" {
				q = "?" + q
			}
			*calls = append(*calls, r.Method+" "+r.URL.Path+q)
		}
		switch r.URL.Path {
		case "/identity/conditionalAccess/policies":
			w.Write([]byte(st.caPolicies))
		case "/identity/conditionalAccess/namedLocations":
			w.Write([]byte(`{"value":[{"@odata.type":"#microsoft.graph.ipNamedLocation","id":"loc1","displayName":"HQ","isTrusted":true}]}`))
		case "/directoryRoles":
			w.Write([]byte(`{"value":[{"id":"r1","displayName":"Global Administrator"}]}`))
		case "/directoryRoles/r1/members":
			w.Write([]byte(st.members))
		case "/policies/authorizationPolicy":
			w.Write([]byte(`{"id":"authorizationPolicy","allowInvitesFrom":"everyone"}`))
		case "/policies/authenticationMethodsPolicy":
			w.WriteHeader(403)
			w.Write([]byte(`{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}`))
		case "/subscribedSkus":
			w.Write([]byte(`{"value":[{"id":"sku1","skuPartNumber":"E3","consumedUnits":10,"prepaidUnits":{"enabled":20}}]}`))
		case "/domains":
			w.Write([]byte(`{"value":[{"id":"contoso.com","isDefault":true}]}`))
		case "/groups":
			w.Write([]byte(`{"value":[{"id":"g2","displayName":"Zeta"},{"id":"g1","displayName":"Alpha"}]}`))
		case "/applications":
			w.Write([]byte(`{"value":[{"id":"a1","appId":"x","displayName":"App","passwordCredentials":[{"keyId":"k1","endDateTime":"2027-01-01T00:00:00Z"}]}]}`))
		case "/servicePrincipals/$count":
			w.Write([]byte("42"))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	})
	sess.SetConfigDir(t.TempDir())
	return NewSnapshotService(sess)
}

func sectionByName(m *SnapshotMeta, name string) SnapshotSection {
	for _, s := range m.Sections {
		if s.Name == name {
			return s
		}
	}
	return SnapshotSection{}
}

func TestSnapshotTakeWritesFileAndSkipsForbiddenSection(t *testing.T) {
	var calls []string
	var events []map[string]any
	eventSink = func(name string, data map[string]any) {
		if name == "snapshot:progress" {
			events = append(events, data)
		}
	}
	t.Cleanup(func() { eventSink = nil })

	svc := snapshotHarness(t, defaultSnapState(), &calls)
	meta, err := svc.Take("Weekly check")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.HasSuffix(meta.ID, "-weekly-check") {
		t.Errorf("id = %q, want <timestamp>-weekly-check", meta.ID)
	}
	if len(meta.Sections) != len(snapshotSections) {
		t.Fatalf("sections = %d, want %d", len(meta.Sections), len(snapshotSections))
	}

	// The forbidden section is skipped with its error; the snapshot still saved.
	amp := sectionByName(meta, "authenticationMethodsPolicy")
	if !amp.Skipped || !strings.Contains(amp.Error, "403") {
		t.Errorf("authenticationMethodsPolicy = %+v, want skipped with a 403", amp)
	}
	if ca := sectionByName(meta, "conditionalAccessPolicies"); ca.Skipped || ca.Count != 2 {
		t.Errorf("conditionalAccessPolicies = %+v", ca)
	}
	if sp := sectionByName(meta, "servicePrincipals"); sp.Count != 1 {
		t.Errorf("servicePrincipals = %+v", sp)
	}

	// Role members are fetched with the narrow $select.
	found := false
	for _, c := range calls {
		if strings.HasPrefix(c, "GET /directoryRoles/r1/members?") && strings.Contains(c, "userPrincipalName") {
			found = true
		}
	}
	if !found {
		t.Errorf("role members not fetched with $select: %v", calls)
	}

	// File on disk, normalized: sorted by id, OData noise gone, @odata.type kept.
	path := filepath.Join(svc.s.ConfigDir(), "snapshots", meta.ID+".json")
	raw, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("snapshot file not written: %v", err)
	}
	var doc snapshotFile
	if err := json.Unmarshal(raw, &doc); err != nil {
		t.Fatal(err)
	}
	groups := doc.Sections["groups"]
	if len(groups) != 2 || groups[0]["id"] != "g1" || groups[1]["id"] != "g2" {
		t.Errorf("groups not sorted by id: %v", groups)
	}
	if strings.Contains(string(raw), "@odata.context") {
		t.Error("@odata.context must be stripped")
	}
	if !strings.Contains(string(raw), "#microsoft.graph.ipNamedLocation") {
		t.Error("@odata.type must be kept (it is the named location kind)")
	}
	if _, ok := doc.Sections["authenticationMethodsPolicy"]; ok {
		t.Error("skipped section must not appear in the data")
	}

	// Progress: one event per section plus the final one.
	if len(events) != len(snapshotSections)+1 {
		t.Errorf("progress events = %d, want %d", len(events), len(snapshotSections)+1)
	}
	if last := events[len(events)-1]; last["done"] != len(snapshotSections) || last["total"] != len(snapshotSections) {
		t.Errorf("final progress event = %v", last)
	}
}

func TestSnapshotDiffDetectsAddedRemovedChanged(t *testing.T) {
	st := defaultSnapState()
	svc := snapshotHarness(t, st, nil)

	a, err := svc.Take("before")
	if err != nil {
		t.Fatal(err)
	}
	// Drift: p1 flips to report-only and widens its exclusions, p2 disappears,
	// p3 appears; a second Global Administrator shows up. modifiedDateTime
	// moves everywhere and must not count as a change on its own.
	st.caPolicies = `{"value":[
		{"id":"p3","displayName":"Block legacy","state":"enabled","modifiedDateTime":"2026-02-01T00:00:00Z"},
		{"id":"p1","displayName":"MFA for all","state":"enabledForReportingButNotEnforced","modifiedDateTime":"2026-02-01T00:00:00Z",
		 "conditions":{"users":{"includeUsers":["All"],"excludeUsers":["u8","u9","u7"]}},
		 "grantControls":{"operator":"OR","builtInControls":["mfa"]}}]}`
	st.members = `{"value":[{"id":"u2","userPrincipalName":"new-admin@contoso.com","displayName":"New"},{"id":"u1","userPrincipalName":"admin@contoso.com","displayName":"Admin"}]}`
	b, err := svc.Take("after")
	if err != nil {
		t.Fatal(err)
	}
	if a.ID == b.ID {
		t.Fatalf("two snapshots must never share an id: %s", a.ID)
	}

	d, err := svc.Diff(a.ID, b.ID)
	if err != nil {
		t.Fatal(err)
	}
	if d.A.ID != a.ID || d.B.ID != b.ID {
		t.Errorf("diff meta = %s → %s", d.A.ID, d.B.ID)
	}
	var ca, roles, amp, groups *SectionDiff
	for i := range d.Sections {
		switch d.Sections[i].Name {
		case "conditionalAccessPolicies":
			ca = &d.Sections[i]
		case "directoryRoles":
			roles = &d.Sections[i]
		case "authenticationMethodsPolicy":
			amp = &d.Sections[i]
		case "groups":
			groups = &d.Sections[i]
		}
	}
	if ca == nil || roles == nil || amp == nil || groups == nil {
		t.Fatalf("sections missing from diff: %+v", d.Sections)
	}

	if len(ca.Added) != 1 || ca.Added[0].Key != "p3" || ca.Added[0].Label != "Block legacy" {
		t.Errorf("CA added = %+v", ca.Added)
	}
	if len(ca.Removed) != 1 || ca.Removed[0].Key != "p2" {
		t.Errorf("CA removed = %+v", ca.Removed)
	}
	if len(ca.Changed) != 1 || ca.Changed[0].Key != "p1" {
		t.Fatalf("CA changed = %+v", ca.Changed)
	}
	paths := map[string]FieldChange{}
	for _, fc := range ca.Changed[0].Changes {
		paths[fc.Path] = fc
	}
	if fc, ok := paths["state"]; !ok || fc.Before != "enabled" || fc.After != "enabledForReportingButNotEnforced" {
		t.Errorf("state change = %+v (all: %v)", fc, paths)
	}
	if _, ok := paths["conditions.users.excludeUsers"]; !ok {
		t.Errorf("excludeUsers change missing: %v", paths)
	}
	if _, ok := paths["modifiedDateTime"]; ok {
		t.Error("modifiedDateTime must be excluded from the diff")
	}
	if len(paths) != 2 {
		t.Errorf("unexpected extra changes: %v", paths)
	}

	if len(roles.Changed) != 1 {
		t.Fatalf("roles changed = %+v", roles.Changed)
	}
	if ch := roles.Changed[0].Changes; len(ch) != 1 || ch[0].Path != "members[u2]" || ch[0].Before != nil {
		t.Errorf("role member change = %+v", ch)
	}
	if !amp.Skipped {
		t.Error("section forbidden on both sides must be reported as skipped, not as empty")
	}
	if groups.Unchanged != 2 || len(groups.Changed)+len(groups.Added)+len(groups.Removed) != 0 {
		t.Errorf("groups should be unchanged: %+v", groups)
	}
	if d.Added != 1 || d.Removed != 1 || d.Changed != 2 {
		t.Errorf("totals = +%d -%d ~%d", d.Added, d.Removed, d.Changed)
	}

	// DiffLatest picks older → newer on its own.
	dl, err := svc.DiffLatest()
	if err != nil {
		t.Fatal(err)
	}
	if dl.A.ID != a.ID || dl.B.ID != b.ID {
		t.Errorf("DiffLatest = %s → %s, want %s → %s", dl.A.ID, dl.B.ID, a.ID, b.ID)
	}
}

func TestSnapshotListGetDelete(t *testing.T) {
	svc := snapshotHarness(t, defaultSnapState(), nil)
	a, _ := svc.Take("first")
	b, _ := svc.Take("second")

	list, err := svc.List()
	if err != nil {
		t.Fatal(err)
	}
	if len(list) != 2 || list[0].ID != b.ID || list[1].ID != a.ID {
		t.Fatalf("list must be newest first: %+v", list)
	}

	raw, err := svc.Get(a.ID)
	if err != nil {
		t.Fatal(err)
	}
	var doc snapshotFile
	if err := json.Unmarshal(raw, &doc); err != nil || doc.Meta.Name != "first" {
		t.Errorf("Get returned %v / %s", err, string(raw)[:60])
	}

	if err := svc.Delete(a.ID); err != nil {
		t.Fatal(err)
	}
	list, _ = svc.List()
	if len(list) != 1 || list[0].ID != b.ID {
		t.Errorf("after delete: %+v", list)
	}
	if _, err := svc.DiffLatest(); err == nil {
		t.Error("DiffLatest with one snapshot must fail")
	}
}

func TestSnapshotRejectsUnsafeIDs(t *testing.T) {
	svc := snapshotHarness(t, defaultSnapState(), nil)
	for _, id := range []string{"", "../secrets", `a\b`, "a/b", ".hidden", "C:x"} {
		if _, err := svc.Get(id); err == nil {
			t.Errorf("Get(%q) must be rejected", id)
		}
		if err := svc.Delete(id); err == nil {
			t.Errorf("Delete(%q) must be rejected", id)
		}
	}
}

func TestSnapshotFailsWhenNothingReadable(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(401)
		w.Write([]byte(`{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}`))
	})
	sess.SetConfigDir(t.TempDir())
	if _, err := NewSnapshotService(sess).Take("x"); err == nil {
		t.Fatal("a snapshot with every section failed must not be saved")
	}
	entries, _ := os.ReadDir(filepath.Join(sess.ConfigDir(), "snapshots"))
	if len(entries) != 0 {
		t.Errorf("no file expected, got %d", len(entries))
	}
}

// diffJSON is the generic engine behind the section diff.
func TestDiffJSONPathsAndKeyedArrays(t *testing.T) {
	parse := func(s string) any {
		var v any
		if err := json.Unmarshal([]byte(s), &v); err != nil {
			t.Fatal(err)
		}
		return v
	}
	a := parse(`{"name":"x","modifiedDateTime":"1","nested":{"flag":true,"list":["a","b"]},
		"items":[{"id":"i1","v":1},{"id":"i2","v":2}],"gone":"yes"}`)
	b := parse(`{"name":"y","modifiedDateTime":"2","nested":{"flag":true,"list":["a","c"]},
		"items":[{"id":"i1","v":10},{"id":"i3","v":3}],"new":"here"}`)

	got := map[string]FieldChange{}
	for _, fc := range diffJSON("", a, b) {
		got[fc.Path] = fc
	}
	want := []string{"name", "nested.list", "items[i1].v", "items[i2]", "items[i3]", "gone", "new"}
	for _, p := range want {
		if _, ok := got[p]; !ok {
			t.Errorf("missing change at %q; got %v", p, got)
		}
	}
	if len(got) != len(want) {
		t.Errorf("unexpected changes: %v", got)
	}
	if fc := got["items[i1].v"]; fc.Before != float64(1) || fc.After != float64(10) {
		t.Errorf("items[i1].v = %+v", fc)
	}
	if fc := got["items[i2]"]; fc.After != nil {
		t.Errorf("removed item must have After=nil: %+v", fc)
	}
	if fc := got["gone"]; fc.Before != "yes" || fc.After != nil {
		t.Errorf("gone = %+v", fc)
	}
	if diffJSON("", a, a) != nil {
		t.Error("identical values must produce no changes")
	}
}

func TestSlugify(t *testing.T) {
	cases := map[string]string{
		"Weekly check":          "weekly-check",
		"  ":                    "snapshot",
		"Before CA change!!!":   "before-ca-change",
		"Снимок до миграции":    "снимок-до-миграции",
		"a/b\\c:d":              "a-b-c-d",
		strings.Repeat("x", 60): strings.Repeat("x", 40),
	}
	for in, want := range cases {
		if got := slugify(in); got != want {
			t.Errorf("slugify(%q) = %q, want %q", in, got, want)
		}
	}
}
