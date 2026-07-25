package journal

import (
	"testing"
)

func TestBeginEventEndRoundTrip(t *testing.T) {
	l := New(t.TempDir())
	l.Begin("op-1", map[string]any{"kind": "transfer", "target": "a → b"})
	l.Event("op-1", "item", map[string]any{"id": "i1", "status": "copied"})
	l.End("op-1", map[string]any{"copied": float64(1)})

	runs := l.List(10)
	if len(runs) != 1 {
		t.Fatalf("want 1 run, got %d", len(runs))
	}
	r := runs[0]
	if r.OpID != "op-1" || r.Kind != "transfer" || r.Target != "a → b" {
		t.Fatalf("bad summary: %+v", r)
	}
	if r.EndedAt == nil || r.Interrupted {
		t.Fatalf("finished run must have EndedAt and not be interrupted: %+v", r)
	}

	full, err := l.Get("op-1")
	if err != nil {
		t.Fatal(err)
	}
	if len(full.Events) != 1 || full.Events[0].Type != "item" {
		t.Fatalf("events: %+v", full.Events)
	}
}

func TestInterruptedRunDetection(t *testing.T) {
	l := New(t.TempDir())
	t.Cleanup(l.Close) // Windows: TempDir cleanup needs the handle released
	l.Begin("op-live", map[string]any{"kind": "playbook", "target": "x"})
	// Still live (no End, writer registered): not interrupted.
	if r := l.List(10); len(r) != 1 || r[0].Interrupted {
		t.Fatalf("live run must not be interrupted: %+v", r)
	}
	// A fresh Log over the same dir (an app restart): no live writers.
	l2 := New(l.dir)
	if r := l2.List(10); len(r) != 1 || !r[0].Interrupted {
		t.Fatalf("run without end record must be interrupted after restart: %+v", r)
	}
}

func TestListNewestFirstAndPrune(t *testing.T) {
	l := New(t.TempDir())
	for _, id := range []string{"a-1", "b-2", "c-3"} { // opIds sort by time by construction
		l.Begin(id, map[string]any{"kind": "transfer", "target": id})
		l.End(id, nil)
	}
	runs := l.List(2)
	if len(runs) != 2 || runs[0].OpID != "c-3" || runs[1].OpID != "b-2" {
		t.Fatalf("want newest first bounded, got %+v", runs)
	}
	l.Prune(1)
	if runs := l.List(10); len(runs) != 1 || runs[0].OpID != "c-3" {
		t.Fatalf("prune must keep the newest: %+v", runs)
	}
}

func TestTornFinalLineIsTolerated(t *testing.T) {
	l := New(t.TempDir())
	l.Begin("op-torn", map[string]any{"kind": "transfer", "target": "t"})
	// Simulate a crash mid-write: append garbage bytes.
	l.mu.Lock()
	f := l.open["op-torn"]
	_, _ = f.WriteString(`{"t":"item","at":`) // torn line
	l.mu.Unlock()
	l.Close() // release the handle (Windows cannot delete open files)

	run, err := New(l.dir).Get("op-torn")
	if err != nil {
		t.Fatal(err)
	}
	if run.Begin == nil || !run.Interrupted {
		t.Fatalf("torn journal must still parse as interrupted: %+v", run)
	}
}
