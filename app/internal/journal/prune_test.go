package journal

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestPruneKeepsInterruptedRunsOutOfTheCount(t *testing.T) {
	dir := t.TempDir()
	l := New(dir)
	// An interrupted (resumable) copy, older than every finished run.
	l.Begin("a-0", map[string]any{"kind": "transfer", "target": "copy"})
	l.Event("a-0", "plan", map[string]any{"driveId": "d"})
	l.Close() // the app exited mid-copy
	l = New(dir)
	for _, id := range []string{"b-1", "c-2", "d-3"} {
		l.Begin(id, map[string]any{"kind": "x", "target": id})
		l.End(id, nil)
	}

	l.Prune(1)
	ids := map[string]bool{}
	for _, r := range l.List(0) {
		ids[r.OpID] = true
	}
	if !ids["a-0"] || !ids["d-3"] || ids["b-1"] || ids["c-2"] || len(ids) != 2 {
		t.Fatalf("after prune: %v — want the interrupted run and the newest finished one", ids)
	}

	// Past the age limit an interrupted run goes too.
	l.prune(1, time.Now().Add(InterruptedMaxAge+time.Hour))
	if _, err := os.Stat(filepath.Join(dir, "a-0.jsonl")); !os.IsNotExist(err) {
		t.Error("a stale interrupted run was kept")
	}
	if _, err := os.Stat(filepath.Join(dir, "d-3.jsonl")); err != nil {
		t.Error("the newest finished run was deleted")
	}
}

func TestHasEndRecordReadsOnlyTheTail(t *testing.T) {
	dir := t.TempDir()
	l := New(dir)
	l.Begin("big-1", map[string]any{"kind": "transfer"})
	for i := 0; i < 5000; i++ { // well past the 64 KB tail
		l.Event("big-1", "item", map[string]any{"id": "item-with-a-longish-identifier", "status": "copied"})
	}
	if hasEndRecord(l.path("big-1")) {
		t.Fatal("no end record yet")
	}
	l.End("big-1", map[string]any{"copied": 5000})
	if !hasEndRecord(l.path("big-1")) {
		t.Fatal("end record not found")
	}

	// An end record longer than the first tail window is still found.
	l.Begin("big-2", map[string]any{"kind": "transfer"})
	l.Event("big-2", "item", map[string]any{"id": "x"})
	l.End("big-2", map[string]any{"note": strings.Repeat("x", 200<<10)})
	if !hasEndRecord(l.path("big-2")) {
		t.Fatal("a long end record was missed")
	}

	// An empty file is not a finished run.
	empty := filepath.Join(dir, "empty-3.jsonl")
	if err := os.WriteFile(empty, nil, 0o644); err != nil {
		t.Fatal(err)
	}
	if hasEndRecord(empty) {
		t.Fatal("an empty file counted as finished")
	}
}
