package services

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func savedSnapshot(t *testing.T, dir string, sections map[string][]map[string]any) string {
	t.Helper()
	doc := &snapshotFile{Sections: sections, Meta: SnapshotMeta{Tenant: "test"}}
	for name, objs := range sections {
		doc.Meta.Sections = append(doc.Meta.Sections, SnapshotSection{Name: name, Count: len(objs)})
	}
	snapDir := filepath.Join(dir, "snapshots")
	if err := os.MkdirAll(snapDir, 0o700); err != nil {
		t.Fatal(err)
	}
	id, err := writeSnapshotExclusive(context.Background(), snapDir, doc, time.Now(), "baseline")
	if err != nil {
		t.Fatal(err)
	}
	return id
}

func TestRestorePlansFieldDiffsAndRecreatesDeleted(t *testing.T) {
	type call struct{ method, path, body string }
	var calls []call
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		calls = append(calls, call{r.Method, r.URL.Path, string(b)})
		switch r.URL.Path {
		case "/identity/conditionalAccess/policies/p1":
			if r.Method == "GET" {
				w.Write([]byte(`{"id":"p1","displayName":"Require MFA","state":"disabled","conditions":{"users":{"includeUsers":["All"]}},"createdDateTime":"2026-01-01T00:00:00Z"}`))
				return
			}
			w.WriteHeader(http.StatusNoContent)
		case "/identity/conditionalAccess/policies/p2":
			w.WriteHeader(http.StatusNotFound)
			w.Write([]byte(`{"error":{"code":"ResourceNotFound","message":"gone"}}`))
		default:
			w.Write([]byte(`{}`))
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"conditionalAccessPolicies": {
			{"id": "p1", "displayName": "Require MFA", "state": "enabled", "conditions": map[string]any{"users": map[string]any{"includeUsers": []any{"All"}}}},
			{"id": "p2", "displayName": "Block legacy auth", "state": "enabled"},
		},
	})
	e := NewEngine(sess)
	p, err := e.Plan("config.restore", map[string]string{"snapshot": id})
	if err != nil {
		t.Fatal(err)
	}
	if len(p.Changes) != 2 || p.ConfirmTarget != "restore 2" {
		t.Fatalf("plan %+v confirm %q", p.Changes, p.ConfirmTarget)
	}
	byTarget := map[string]string{}
	for _, c := range p.Changes {
		byTarget[c.Target] = c.Op + ":" + c.Before
	}
	if byTarget["Require MFA"] != "set:state" || byTarget["Block legacy auth"] != "add:" {
		t.Fatalf("changes %v", byTarget)
	}
	if _, err := e.Apply(p.ID, "restore 2"); err != nil {
		t.Fatal(err)
	}
	var patched, posted string
	for _, c := range calls {
		if c.method == "PATCH" {
			patched = c.body
		}
		if c.method == "POST" && c.path == "/identity/conditionalAccess/policies" {
			posted = c.body
		}
	}
	var body map[string]any
	_ = json.Unmarshal([]byte(patched), &body)
	if body["state"] != "enabled" || body["createdDateTime"] != nil || body["id"] != nil {
		t.Fatalf("patch body %s (only writable fields, snapshot values)", patched)
	}
	if !strings.Contains(posted, "Block legacy auth") {
		t.Fatalf("deleted policy must be recreated: %s", posted)
	}
}

func TestExportWritesOneStableFilePerObjectWithoutSecrets(t *testing.T) {
	doc := &snapshotFile{
		Meta: SnapshotMeta{ID: "20261008-120000-baseline", Name: "baseline"},
		Sections: map[string][]map[string]any{
			"applications": {
				{"id": "a1", "displayName": "Payroll", "passwordCredentials": []any{map[string]any{"hint": "ab", "secretText": "SHOULD-NOT-LEAK"}}},
				{"id": "a2", "displayName": "Payroll"},
			},
		},
	}
	root := t.TempDir()
	dir, err := writeExport(doc, root)
	if err != nil {
		t.Fatal(err)
	}
	entries, _ := os.ReadDir(filepath.Join(dir, "applications"))
	if len(entries) != 2 {
		t.Fatalf("files %v (same-named objects must not overwrite each other)", entries)
	}
	for _, e := range entries {
		b, _ := os.ReadFile(filepath.Join(dir, "applications", e.Name()))
		if strings.Contains(string(b), "SHOULD-NOT-LEAK") {
			t.Fatalf("secret exported in %s", e.Name())
		}
	}
	if _, err := os.Stat(filepath.Join(dir, "snapshot.json")); err != nil {
		t.Fatal(err)
	}
	if _, err := writeExport(&snapshotFile{Meta: SnapshotMeta{ID: "../evil"}}, root); err == nil {
		t.Fatal("a path-like id must be refused")
	}
}

func TestRestoreRefusesAnotherTenantAndRemapsLocations(t *testing.T) {
	var posted string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/identity/conditionalAccess/namedLocations" && r.Method == "GET":
			// The office location was recreated after deletion: new id.
			w.Write([]byte(`{"value":[{"id":"loc-new","displayName":"Office"}]}`))
		case r.URL.Path == "/identity/conditionalAccess/policies" && r.Method == "GET":
			w.Write([]byte(`{"value":[]}`))
		case r.URL.Path == "/identity/conditionalAccess/policies" && r.Method == "POST":
			b, _ := io.ReadAll(r.Body)
			posted = string(b)
			w.Write([]byte(`{"id":"p9"}`))
		default:
			w.WriteHeader(http.StatusNotFound)
			w.Write([]byte(`{"error":{"code":"ResourceNotFound","message":"gone"}}`))
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sections := map[string][]map[string]any{
		"namedLocations": {{"id": "loc-old", "displayName": "Office"}},
		"conditionalAccessPolicies": {{"id": "p1", "displayName": "Office only", "state": "enabled",
			"conditions":    map[string]any{"locations": map[string]any{"includeLocations": []any{"All"}, "excludeLocations": []any{"loc-old"}}},
			"grantControls": map[string]any{"operator": "OR", "authenticationStrength": map[string]any{"id": "s1", "displayName": "MFA", "createdDateTime": "2026-01-01T00:00:00Z"}},
		}},
	}
	id := savedSnapshot(t, dir, sections)
	e := NewEngine(sess)
	p, err := e.Plan("config.restore", map[string]string{"snapshot": id})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(posted, `"loc-new"`) || strings.Contains(posted, "loc-old") {
		t.Fatalf("location ids must be remapped by name: %s", posted)
	}
	if strings.Contains(posted, "createdDateTime") {
		t.Fatalf("authentication strength must be sent by id only: %s", posted)
	}

	// A location that cannot be resolved stops the plan.
	sections["namedLocations"] = nil
	id2 := savedSnapshot(t, dir, sections)
	if _, err := e.Plan("config.restore", map[string]string{"snapshot": id2}); err == nil || !strings.Contains(err.Error(), "loc-old") {
		t.Fatalf("missing location must be refused: %v", err)
	}

	// Another tenant's snapshot is never restored.
	sess.SetIdentity("tenant-b", false)
	other := &snapshotFile{Meta: SnapshotMeta{Tenant: "test", TenantID: "tenant-a"}, Sections: sections}
	oid, err := writeSnapshotExclusive(context.Background(), filepath.Join(dir, "snapshots"), other, time.Now().Add(time.Minute), "a")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := e.Plan("config.restore", map[string]string{"snapshot": oid}); err == nil || !strings.Contains(err.Error(), "another tenant") {
		t.Fatalf("cross-tenant restore must be refused: %v", err)
	}
}

func TestExportStripsSecretLookingFields(t *testing.T) {
	out := stripSecrets(map[string]any{
		"wifi":     map[string]any{"ssid": "corp", "preSharedKey": "x1"},
		"cred":     map[string]any{"hint": "abc", "displayName": "ci"},
		"oma":      []any{map[string]any{"isEncrypted": true, "value": "x2", "omaUri": "./a"}, map[string]any{"isEncrypted": false, "value": "plain"}},
		"vpnCreds": map[string]any{"clientSecret": "x3", "Passphrase": "x4"},
	})
	b, _ := json.Marshal(out)
	for _, leak := range []string{"x1", "abc", "x2", "x3", "x4"} {
		if strings.Contains(string(b), `"`+leak+`"`) {
			t.Fatalf("%s leaked: %s", leak, b)
		}
	}
	if !strings.Contains(string(b), "plain") || !strings.Contains(string(b), "corp") {
		t.Fatalf("ordinary values must stay: %s", b)
	}
}
