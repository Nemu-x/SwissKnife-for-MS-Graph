package services

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/pwsh"
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
	// The list still shows; the problem is reported separately.
	if list, err := svc.List(); err != nil || list == nil {
		t.Fatalf("List = %v, %v; want the packs without an error", list, err)
	}
	copies := corruptCopies(t, dir, "packs.json")
	if len(copies) != 1 {
		t.Fatalf("set-aside copies = %v", copies)
	}
	if b, _ := os.ReadFile(copies[0]); string(b) != `{"pinned": {"x": ` {
		t.Errorf("copy = %q", b)
	}
	if err := svc.TrustProblem(); err == nil || !strings.Contains(err.Error(), "packs.json.corrupt-") {
		t.Fatalf("TrustProblem = %v, want the set-aside error", err)
	}
	if err := svc.TrustProblem(); err != nil {
		t.Errorf("second TrustProblem = %v, want the problem reported only once", err)
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

func TestTrustChangeOnACorruptStoreReportsFirst(t *testing.T) {
	dir := t.TempDir()
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(dir)
	path := filepath.Join(dir, "packs.json")
	if err := os.WriteFile(path, []byte(`not json`), 0o600); err != nil {
		t.Fatal(err)
	}
	svc := NewPacksService(sess)
	if _, err := svc.SetHub(""); err == nil || !strings.Contains(err.Error(), "packs.json") {
		t.Fatalf("SetHub on a corrupt store = %v, want the set-aside report", err)
	}
	if _, err := os.Stat(path); !os.IsNotExist(err) {
		t.Error("the refused change wrote packs.json")
	}
	if _, err := svc.SetHub(""); err != nil {
		t.Errorf("after the report the change goes through: %v", err)
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
	// A PowerShell backend that needs the worker says why it is unavailable
	// (remoteFor sets the file aside) until Settings reports it.
	p := workerAware{inner: unavailablePS{}, family: pwsh.FamilyExchange}
	if r := p.Status(sess); r == nil || r.Key != "workerFileCorrupt" {
		t.Fatalf("worker-aware status = %+v, want workerFileCorrupt", r)
	}
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
