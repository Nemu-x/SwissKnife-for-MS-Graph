package services

import (
	"fmt"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/session"
)

// TestResumeCopyContinuesInterruptedRun: a journaled server-side copy that
// died mid-run must resume — completed items counted, the in-flight monitor
// re-polled, pending items issued — and finally get its terminal record.
func TestResumeCopyContinuesInterruptedRun(t *testing.T) {
	var copyPosts []string
	var srvURL string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "POST" && strings.Contains(r.URL.Path, "/copy"):
			copyPosts = append(copyPosts, r.URL.Path)
			w.Header().Set("Location", srvURL+"/monitor-new")
			w.WriteHeader(http.StatusAccepted)
		case r.URL.Path == "/monitor-old" || r.URL.Path == "/monitor-new":
			fmt.Fprint(w, `{"status":"completed","percentageComplete":100}`)
		default:
			w.Write([]byte(`{}`))
		}
	}))
	t.Cleanup(srv.Close)
	srvURL = srv.URL

	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	dir := t.TempDir()
	j := journal.New(dir)

	// Journal of a run that died: i1 finished, i2 was in flight, i3 never started.
	j.Begin("run-x", map[string]any{"kind": "transfer", "target": "alice → arch", "source": "alice@contoso.com", "dest": "arch@contoso.com"})
	j.Event("run-x", "plan", map[string]any{
		"driveId": "d2", "parentId": "fld1", "behavior": "fail", "sourceUser": "alice@contoso.com",
		"items": []any{
			map[string]any{"id": "i1", "name": "Docs", "size": float64(100)},
			map[string]any{"id": "i2", "name": "a.txt", "size": float64(50)},
			map[string]any{"id": "i3", "name": "b.txt", "size": float64(25)},
		},
	})
	j.Event("run-x", "item", map[string]any{"id": "i1", "status": "copied"})
	j.Event("run-x", "item", map[string]any{"id": "i2", "status": "inflight", "monitor": srvURL + "/monitor-old"})
	// No end record. Reopening the same directory with a fresh Log (no live
	// writers) simulates the app restart — the run must read as interrupted.
	j.Close()
	restarted := journal.New(dir)
	sess.SetJournal(restarted)

	drive := NewDriveService(sess)
	res, err := drive.ResumeCopy("run-x")
	if err != nil {
		t.Fatal(err)
	}
	// i1 counted from the previous run, i2 completed via the old monitor,
	// i3 issued fresh — 3 copied total, exactly one new copy POST (i3).
	if len(res.Copied) != 3 {
		t.Fatalf("want 3 copied, got %v (failed=%v)", res.Copied, res.Failed)
	}
	if len(copyPosts) != 1 || !strings.Contains(copyPosts[0], "/items/i3/") {
		t.Fatalf("only i3 must be re-issued, got %v", copyPosts)
	}

	run, err := restarted.Get("run-x")
	if err != nil {
		t.Fatal(err)
	}
	if run.EndedAt == nil {
		t.Fatal("resumed run must receive a terminal record")
	}
	if run.Summary["copied"] != float64(3) {
		t.Fatalf("terminal summary wrong: %+v", run.Summary)
	}
}

// TestResumeCopyRejectsFinishedRuns: a run with a terminal record is not resumable.
func TestResumeCopyRejectsFinishedRuns(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) })
	j := journal.New(t.TempDir())
	sess.SetJournal(j)
	j.Begin("run-done", map[string]any{"kind": "transfer", "target": "x"})
	j.End("run-done", map[string]any{"copied": 1})

	if _, err := NewDriveService(sess).ResumeCopy("run-done"); err == nil {
		t.Fatal("finished runs must not resume")
	}
}
