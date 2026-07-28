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
		switch {
		case r.URL.Path == "/users/lead@contoso.com":
			w.Write([]byte(`{"id":"new-owner-id"}`))
		case r.URL.Path == "/users/leaver@contoso.com":
			w.Write([]byte(`{"id":"leaver-id","displayName":"Leaver"}`))
		case r.URL.Path == "/users/leaver@contoso.com/ownedObjects":
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

// Meetings the leaver merely attends belong to somebody else and must survive;
// only the ones they organise are cancelled, and only with attendees notified.
func TestOffboardCancelsOnlyOrganisedFutureMeetings(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch {
		case r.URL.Path == "/users/leaver@contoso.com/events":
			if f := r.URL.Query().Get("$filter"); !strings.Contains(f, "start/dateTime ge") {
				t.Errorf("future-only filter missing, got %q", f)
			}
			w.Write([]byte(`{"value":[
				{"id":"e1","isOrganizer":true,"isCancelled":false,"attendees":[{"type":"required"}]},
				{"id":"e2","isOrganizer":true,"isCancelled":false,"attendees":[]},
				{"id":"e3","isOrganizer":false,"isCancelled":false,"attendees":[{"type":"required"}]},
				{"id":"e4","isOrganizer":true,"isCancelled":true,"attendees":[]}]}`))
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
	for _, bad := range []string{"events/e3", "events/e4"} {
		if strings.Contains(joined, bad) {
			t.Errorf("must not touch %s (not organised / already cancelled):\n%s", bad, joined)
		}
	}
}
