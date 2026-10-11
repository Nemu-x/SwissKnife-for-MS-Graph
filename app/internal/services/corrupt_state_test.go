package services

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/session"
)

func corruptCopies(t *testing.T, dir, name string) []string {
	t.Helper()
	m, _ := filepath.Glob(filepath.Join(dir, name+".corrupt-*"))
	return m
}

func TestCorruptPacksJSONIsSetAsideAndReportedOnce(t *testing.T) {
	dir := t.TempDir()
	logDir := t.TempDir()
	sess := session.New(auditlog.New(logDir))
	sess.SetConfigDir(dir)
	path := filepath.Join(dir, "packs.json")
	if err := os.WriteFile(path, []byte(`{"pinned": {"x": `), 0o600); err != nil {
		t.Fatal(err)
	}
	svc := NewPacksService(sess)
	_, err := svc.List()
	if err == nil || !strings.Contains(err.Error(), "packs.json.corrupt-") {
		t.Fatalf("List = %v, want the set-aside error", err)
	}
	copies := corruptCopies(t, dir, "packs.json")
	if len(copies) != 1 {
		t.Fatalf("set-aside copies = %v", copies)
	}
	if b, _ := os.ReadFile(copies[0]); string(b) != `{"pinned": {"x": ` {
		t.Errorf("copy = %q", b)
	}
	if _, err := svc.List(); err != nil {
		t.Errorf("second List = %v, want the problem reported only once", err)
	}
	// The next save writes a fresh file; the set-aside copy stays.
	if _, err := svc.Untrust("anything"); err != nil {
		t.Fatal(err)
	}
	if len(corruptCopies(t, dir, "packs.json")) != 1 {
		t.Error("the set-aside copy was lost")
	}
	if b, _ := os.ReadFile(filepath.Join(logDir, "actions.log.jsonl")); !strings.Contains(string(b), "pack.trustCorrupt") {
		t.Errorf("no audit entry: %s", b)
	}
}

func TestCorruptWorkerJSONIsReportedByStatus(t *testing.T) {
	dir := t.TempDir()
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(dir)
	if err := os.WriteFile(filepath.Join(dir, "worker.json"), []byte("not json"), 0o600); err != nil {
		t.Fatal(err)
	}
	ws := NewWorkerService(sess)
	if w, err := ws.Status(); err == nil || w != nil || !strings.Contains(err.Error(), "worker.json") {
		t.Fatalf("Status = %v, %v; want the set-aside error", w, err)
	}
	if len(corruptCopies(t, dir, "worker.json")) != 1 {
		t.Error("worker.json was not set aside")
	}
	if w, err := ws.Status(); err != nil || w != nil {
		t.Errorf("second Status = %v, %v; want unpaired, no error", w, err)
	}
}
