package services

import (
	"encoding/json"
	"net/http"
	"strings"
	"testing"
)

// TestCreateUniqueRootNeverReusesLastRun: a root folder that already exists
// (a previous, possibly partial run) must get a numeric suffix. Reusing it and
// importing every item again in create mode would duplicate what already made
// it across.
func TestCreateUniqueRootNeverReusesLastRun(t *testing.T) {
	var names []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.Method != "POST" || r.URL.Path != "/admin/exchange/mailboxes/MBX:t/folders" {
			t.Errorf("unexpected call %s %s", r.Method, r.URL.Path)
			w.WriteHeader(500)
			return
		}
		var body map[string]any
		_ = json.NewDecoder(r.Body).Decode(&body)
		name, _ := body["displayName"].(string)
		names = append(names, name)
		if len(names) <= 2 {
			// First two candidates already exist from earlier runs.
			w.WriteHeader(409)
			_, _ = w.Write([]byte(`{"error":{"code":"ErrorFolderExists","message":"A folder with this name already exists."}}`))
			return
		}
		_, _ = w.Write([]byte(`{"id":"f-new"}`))
	})
	c, err := sess.Client()
	if err != nil {
		t.Fatal(err)
	}
	svc := NewMailboxTransferService(sess)
	id, name, err := svc.createUniqueRoot(sess.Ctx(), c, "MBX:t", "Archive - Alice")
	if err != nil {
		t.Fatal(err)
	}
	if id != "f-new" || name != "Archive - Alice (3)" {
		t.Errorf("got id=%q name=%q, want f-new / Archive - Alice (3)", id, name)
	}
	if strings.Join(names, ";") != "Archive - Alice;Archive - Alice (2);Archive - Alice (3)" {
		t.Errorf("candidate order = %v", names)
	}
}

// A non-conflict error must surface unchanged, not be retried with suffixes.
func TestCreateUniqueRootStopsOnRealError(t *testing.T) {
	calls := 0
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls++
		w.WriteHeader(403)
		_, _ = w.Write([]byte(`{"error":{"code":"ErrorAccessDenied","message":"no"}}`))
	})
	c, _ := sess.Client()
	if _, _, err := NewMailboxTransferService(sess).createUniqueRoot(sess.Ctx(), c, "MBX:t", "Archive"); err == nil {
		t.Fatal("expected the 403 to be returned")
	}
	if calls != 1 {
		t.Errorf("a 403 must not be retried, got %d calls", calls)
	}
}
