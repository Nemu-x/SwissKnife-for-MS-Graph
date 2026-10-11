package services

import (
	"net/http"
	"strings"
	"testing"
)

func TestDriftAlertsOncePerNewDifference(t *testing.T) {
	denyCA := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == "/identity/conditionalAccess/policies" && denyCA {
			w.WriteHeader(http.StatusForbidden)
			w.Write([]byte(`{"error":{"code":"Authorization_RequestDenied","message":"no"}}`))
			return
		}
		if r.URL.Path == "/identity/conditionalAccess/policies" {
			w.Write([]byte(`{"value":[{"id":"p1","displayName":"Require MFA","state":"disabled"}]}`))
			return
		}
		w.Write([]byte(`{"value":[]}`))
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sess.SetIdentity("tenant-a", true)
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"conditionalAccessPolicies": {{"id": "p1", "displayName": "Require MFA", "state": "enabled"}},
	})
	var events []map[string]any
	SetEventSink(func(name string, data map[string]any) {
		if name == "drift:detected" {
			events = append(events, data)
		}
	})
	t.Cleanup(func() { SetEventSink(nil) })

	x := NewSnapshotService(sess)
	if _, err := x.SetDriftWatch(DriftWatch{BaselineID: id, EveryHours: 7}); err == nil {
		t.Fatal("only 0, 1, 6 or 24 hours are allowed")
	}
	if _, err := x.SetDriftWatch(DriftWatch{BaselineID: id, EveryHours: 6}); err != nil {
		t.Fatal(err)
	}
	sum, err := x.checkDrift(false)
	if err != nil {
		t.Fatal(err)
	}
	if sum.Changed != 1 || !strings.Contains(strings.Join(sum.Sections, ","), "conditionalAccessPolicies") {
		t.Fatalf("summary %+v", sum)
	}
	if len(events) != 1 {
		t.Fatalf("events %d, want 1", len(events))
	}
	// The same drift again: no second alert from the schedule…
	if _, err := x.checkDrift(false); err != nil {
		t.Fatal(err)
	}
	if len(events) != 1 {
		t.Fatalf("the same drift must alert once, got %d", len(events))
	}
	// …but a manual check still reports it.
	if _, err := x.CheckDriftNow(); err != nil {
		t.Fatal(err)
	}
	if len(events) != 2 {
		t.Fatalf("a manual check must report: %d", len(events))
	}
	w, _ := x.GetDriftWatch()
	if w.LastSummary == nil || w.LastCheck.IsZero() {
		t.Fatalf("watch %+v", w)
	}
	// A section that briefly cannot be read keeps its drift: no new alerts
	// when it fails, nor when it reads again.
	denyCA = true
	if _, err := x.checkDrift(false); err != nil {
		t.Fatal(err)
	}
	denyCA = false
	if _, err := x.checkDrift(false); err != nil {
		t.Fatal(err)
	}
	if len(events) != 2 {
		t.Fatalf("a flaky section must not re-alert: %d", len(events))
	}
}
