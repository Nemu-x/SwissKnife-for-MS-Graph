package services

import (
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// harness builds a session wired to a mock Graph server. The handler receives
// every request so tests can assert paths/bodies and shape responses.
func harness(t *testing.T, handler http.HandlerFunc) *session.Session {
	t.Helper()
	srv := httptest.NewServer(handler)
	t.Cleanup(srv.Close)

	audit := auditlog.New(t.TempDir())
	sess := session.New(audit)
	client := graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL))
	sess.SetClient(client, "test")
	return sess
}

// TestServerSideCopyCopiesTopLevelItems: the copy must run via Graph's async
// /copy operation (bytes never transit the operator machine) — no /content
// download or upload endpoints may be touched.
func TestServerSideCopyCopiesTopLevelItems(t *testing.T) {
	var calls []string
	var srvURL string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch {
		case r.Method == "GET" && r.URL.Path == "/users/arch@contoso.com/drive":
			w.Write([]byte(`{"id":"d2"}`))
		case r.Method == "POST" && r.URL.Path == "/users/arch@contoso.com/drive/root/children":
			w.Write([]byte(`{"id":"newf"}`))
		case r.Method == "GET" && r.URL.Path == "/users/arch@contoso.com/drive/root:/backup":
			w.Write([]byte(`{"id":"fld1"}`))
		case r.Method == "GET" && r.URL.Path == "/users/alice@contoso.com/drive/root/children":
			w.Write([]byte(`{"value":[
				{"id":"i1","name":"Docs","size":100,"folder":{}},
				{"id":"i2","name":"a.txt","size":50}]}`))
		case r.Method == "POST" && strings.HasPrefix(r.URL.Path, "/users/alice@contoso.com/drive/items/"):
			if got := r.URL.Query().Get("@microsoft.graph.conflictBehavior"); got != "fail" {
				t.Errorf("conflictBehavior: want fail, got %q", got)
			}
			w.Header().Set("Location", srvURL+"/monitor")
			w.WriteHeader(http.StatusAccepted)
		case r.Method == "GET" && r.URL.Path == "/monitor":
			w.Write([]byte(`{"status":"completed","percentageComplete":100}`))
		default:
			w.Write([]byte(`{}`))
		}
	}))
	t.Cleanup(srv.Close)
	srvURL = srv.URL

	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	drive := NewDriveService(sess)

	res, err := drive.CopyBetweenUsers("alice@contoso.com", "arch@contoso.com", "backup", false)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Copied) != 2 || res.Copied[0] != "Docs/" || res.Copied[1] != "a.txt" {
		t.Fatalf("copied: %v", res.Copied)
	}
	if len(res.Failed) != 0 || res.Canceled {
		t.Fatalf("unexpected failures/cancel: %+v", res)
	}
	for _, c := range calls {
		if strings.Contains(c, "/content") {
			t.Fatalf("bytes must not transit locally, saw %q in %v", c, calls)
		}
	}
}

func TestOffboardCancelSkipsRemainingSteps(t *testing.T) {
	var calls []string
	var pb *PlaybookService
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		// Call 1 is the identity prefetch; the operator hits Cancel while the
		// first STEP (call 2) is executing.
		if len(calls) == 2 {
			pb.Cancel()
		}
		w.Write([]byte(`{}`))
	})
	pb = NewPlaybookService(sess)

	res, err := pb.Offboard(fullOffboardRequest())
	if err != nil {
		t.Fatal(err)
	}
	if !res.Canceled {
		t.Fatal("result must be marked canceled")
	}
	if len(res.Steps) != 1 {
		t.Fatalf("only the in-flight step should be reported, got %+v", res.Steps)
	}
	if len(calls) != 2 {
		t.Fatalf("no further Graph calls after cancel, got %v", calls)
	}
}

func TestUsersListBuildsFilter(t *testing.T) {
	var gotQuery string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		gotQuery = r.URL.RawQuery
		w.Write([]byte(`{"value":[{"id":"1","displayName":"Alice"}]}`))
	})
	users := NewUsersService(sess)

	out, err := users.List("ali", 0)
	if err != nil {
		t.Fatal(err)
	}
	if len(out) != 1 {
		t.Fatalf("want 1 user, got %d", len(out))
	}
	if !strings.Contains(gotQuery, "startswith") || !strings.Contains(gotQuery, "ali") {
		t.Errorf("filter not applied: %s", gotQuery)
	}
}

func TestReadOnlyBlocksWrite(t *testing.T) {
	called := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		called = true
		w.WriteHeader(200)
	})
	sess.SetReadOnly(true)
	users := NewUsersService(sess)

	err := users.Block("bob@contoso.com")
	if err != session.ErrReadOnly {
		t.Fatalf("want ErrReadOnly, got %v", err)
	}
	if called {
		t.Error("write reached the server despite read-only mode")
	}
}

func TestResetPasswordRequiresTypedConfirm(t *testing.T) {
	patched := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.Method == http.MethodPatch {
			patched = true
		}
		w.WriteHeader(204)
	})
	users := NewUsersService(sess)

	// wrong confirm -> blocked, no PATCH
	if err := users.ResetPassword("alice@contoso.com", "P@ss", true, "wrong"); err == nil {
		t.Fatal("want error on confirm mismatch")
	}
	if patched {
		t.Fatal("PATCH sent despite confirm mismatch")
	}
	// correct confirm -> proceeds
	if err := users.ResetPassword("alice@contoso.com", "P@ss", true, "alice@contoso.com"); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !patched {
		t.Error("PATCH not sent with correct confirm")
	}
}

func TestIntuneWipeConfirmMismatch(t *testing.T) {
	posted := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		posted = true
		w.WriteHeader(204)
	})
	intune := NewIntuneService(sess)

	if err := intune.Wipe("device-1", false, false, "device-WRONG"); err == nil {
		t.Fatal("want error on confirm mismatch")
	}
	if posted {
		t.Error("wipe POST sent despite confirm mismatch")
	}
}

func TestTeamsCreatePrivateChannelNeedsOwner(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{}`))
	})
	teams := NewTeamsService(sess)

	if _, err := teams.CreateChannel("team-1", "Secret", "", "private", ""); err == nil {
		t.Fatal("want error: private channel without owner")
	}
}

func TestOffboardingPreviewCountsRecursively(t *testing.T) {
	// root: file a (10), folder F; F children: file b (20), file c (30)
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.Contains(r.URL.Path, "/drive/root/children"):
			w.Write([]byte(`{"value":[{"id":"a","name":"a.txt","size":10},{"id":"F","name":"F","folder":{}}]}`))
		case strings.Contains(r.URL.Path, "/drive/items/F/children"):
			w.Write([]byte(`{"value":[{"id":"b","name":"b.txt","size":20},{"id":"c","name":"c.txt","size":30}]}`))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	})
	drive := NewDriveService(sess)

	prev, err := drive.OffboardingPreview("alice@contoso.com")
	if err != nil {
		t.Fatal(err)
	}
	if prev.Files != 3 || prev.Folders != 1 || prev.TotalBytes != 60 {
		t.Errorf("preview = %+v, want files=3 folders=1 bytes=60", prev)
	}
}

func TestRawSendRejectsBadMethod(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) })
	raw := NewRawService(sess)
	if _, err := raw.Send("FETCH", "/me", ""); err == nil {
		t.Fatal("want error for unsupported method")
	}
}

func TestRawSendRejectsInvalidJSONBody(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) })
	raw := NewRawService(sess)
	if _, err := raw.Send("POST", "/groups", "{not json"); err == nil {
		t.Fatal("want error for invalid JSON body")
	}
}

func TestLicensingAssignBuildsBody(t *testing.T) {
	var body map[string]any
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		json.NewDecoder(r.Body).Decode(&body)
		w.Write([]byte(`{}`))
	})
	lic := NewLicensingService(sess)
	if _, err := lic.Assign("alice@contoso.com", []string{"sku-1"}, nil); err != nil {
		t.Fatal(err)
	}
	add, _ := body["addLicenses"].([]any)
	if len(add) != 1 {
		t.Errorf("addLicenses = %v", body["addLicenses"])
	}
	if _, ok := body["removeLicenses"]; !ok {
		t.Error("removeLicenses missing (should be [] not null)")
	}
}
