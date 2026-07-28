package services

import (
	"net/http"
	"strings"
	"testing"
)

// Ownership transfer must add the new owner BEFORE dropping the leaver: a group
// with no owner at all is exactly what this step exists to prevent.
func TestOffboardTransfersGroupOwnershipInSafeOrder(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch r.URL.Path {
		case "/users/lead@contoso.com":
			w.Write([]byte(`{"id":"new-owner-id"}`))
		case "/users/leaver@contoso.com":
			w.Write([]byte(`{"id":"leaver-id","displayName":"Leaver"}`))
		case "/users/leaver@contoso.com/ownedObjects":
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.group","id":"g1","displayName":"Global Finance"},
				{"@odata.type":"#microsoft.graph.application","id":"a1","displayName":"Some app"}]}`))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	})

	res, err := NewPlaybookService(sess).Offboard(OffboardRequest{
		Upn: "leaver@contoso.com", Confirm: "leaver@contoso.com",
		TransferOwnershipTo: "lead@contoso.com",
	})
	if err != nil {
		t.Fatal(err)
	}

	addIdx, delIdx := -1, -1
	for i, c := range calls {
		if c == "POST /groups/g1/owners/$ref" {
			addIdx = i
		}
		if c == "DELETE /groups/g1/owners/leaver-id/$ref" {
			delIdx = i
		}
		// Owned applications are deliberately left alone.
		if strings.Contains(c, "/applications/a1/owners") {
			t.Errorf("application ownership must not be touched: %s", c)
		}
	}
	if addIdx < 0 || delIdx < 0 {
		t.Fatalf("expected an owner add and an owner removal, got:\n%s", strings.Join(calls, "\n"))
	}
	if addIdx > delIdx {
		t.Errorf("the new owner must be added before the leaver is removed:\n%s", strings.Join(calls, "\n"))
	}
	if len(res.Steps) == 0 {
		t.Error("the transfer must be reported as a step")
	}
}

// The new owner may already own the group: Graph rejects the duplicate
// reference, and that rejection must not abort the hand-over.
func TestOffboardOwnershipSurvivesAnExistingOwner(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch {
		case r.Method == "POST" && r.URL.Path == "/groups/g1/owners/$ref":
			w.WriteHeader(http.StatusBadRequest)
			w.Write([]byte(`{"error":{"code":"Request_BadRequest","message":"One or more added object references already exist for the following modified properties: 'owners'."}}`))
		case r.URL.Path == "/users/lead@contoso.com":
			w.Write([]byte(`{"id":"new-owner-id"}`))
		case r.URL.Path == "/users/leaver@contoso.com":
			w.Write([]byte(`{"id":"leaver-id"}`))
		case r.URL.Path == "/users/leaver@contoso.com/ownedObjects":
			w.Write([]byte(`{"value":[{"@odata.type":"#microsoft.graph.group","id":"g1","displayName":"Global Finance"}]}`))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	})

	res, err := NewPlaybookService(sess).Offboard(OffboardRequest{
		Upn: "leaver@contoso.com", Confirm: "leaver@contoso.com", TransferOwnershipTo: "lead@contoso.com",
	})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(strings.Join(calls, "\n"), "DELETE /groups/g1/owners/leaver-id/$ref") {
		t.Errorf("an already-owner must not block dropping the leaver:\n%s", strings.Join(calls, "\n"))
	}
	// A vacuous loop would pass when no step was emitted at all.
	transferred := false
	for _, s := range res.Steps {
		if s.Name != "Transfer ownership" {
			continue
		}
		transferred = true
		if !s.OK {
			t.Errorf("the step must succeed, got error %q", s.Error)
		}
	}
	if !transferred {
		t.Errorf("no ownership transfer step was reported: %+v", res.Steps)
	}
}

// Meetings the leaver merely attends belong to somebody else and must survive;
// only the ones they organise are cancelled, and only with attendees notified.
func TestOffboardCancelsOnlyOrganisedFutureMeetings(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch r.URL.Path {
		case "/users/leaver@contoso.com/events":
			if f := r.URL.Query().Get("$filter"); !strings.Contains(f, "start/dateTime ge") {
				t.Errorf("future-only filter missing, got %q", f)
			}
			w.Write([]byte(`{"value":[
				{"id":"e1","isOrganizer":true,"isCancelled":false,"attendees":[{"type":"required"}]},
				{"id":"e2","isOrganizer":true,"isCancelled":false,"attendees":[]},
				{"id":"e3","isOrganizer":false,"isCancelled":false,"attendees":[{"type":"required"}]},
				{"id":"e4","isOrganizer":true,"isCancelled":true,"attendees":[]},
				{"id":"e5","type":"seriesMaster","isOrganizer":true,"isCancelled":false,"attendees":[{"type":"required"}],
				 "recurrence":{"range":{"type":"noEnd"}}},
				{"id":"e6","type":"seriesMaster","isOrganizer":true,"isCancelled":false,"attendees":[{"type":"required"}],
				 "recurrence":{"range":{"type":"endDate","endDate":"2020-01-01"}}}]}`))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	})

	if _, err := NewPlaybookService(sess).Offboard(OffboardRequest{
		Upn: "leaver@contoso.com", Confirm: "leaver@contoso.com", CancelFutureEvents: true,
	}); err != nil {
		t.Fatal(err)
	}

	joined := strings.Join(calls, "\n")
	// With attendees: cancel (they get a notification).
	if !strings.Contains(joined, "POST /users/leaver@contoso.com/events/e1/cancel") {
		t.Errorf("a meeting with attendees must be cancelled, not deleted:\n%s", joined)
	}
	// Without attendees: a plain delete is enough.
	if !strings.Contains(joined, "DELETE /users/leaver@contoso.com/events/e2") {
		t.Errorf("a solo appointment should just be deleted:\n%s", joined)
	}
	// A recurring series still running must be cancelled even though its master
	// started in the past; one that already ended must be left alone.
	if !strings.Contains(joined, "POST /users/leaver@contoso.com/events/e5/cancel") {
		t.Errorf("a live recurring series must be cancelled:\n%s", joined)
	}
	for _, bad := range []string{"events/e3", "events/e4", "events/e6"} {
		if strings.Contains(joined, bad) {
			t.Errorf("must not touch %s (not organised / already cancelled / series ended):\n%s", bad, joined)
		}
	}
}
