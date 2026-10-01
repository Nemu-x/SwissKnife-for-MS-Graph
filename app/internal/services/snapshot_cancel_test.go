package services

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"testing"
)

// A cancel that lands with the final progress event (every section already
// collected) must still abort Take: no error swallowed, no file on disk.
func TestSnapshotTakeHonoursLateCancel(t *testing.T) {
	svc := snapshotHarness(t, defaultSnapState(), nil)
	SetEventSink(func(name string, data map[string]any) {
		if name != "snapshot:progress" {
			return
		}
		done, _ := data["done"].(int)
		total, _ := data["total"].(int)
		if data["section"] == "" && done == total {
			svc.Cancel() // the operator presses Cancel just as collection ends
		}
	})
	t.Cleanup(func() { SetEventSink(nil) })

	meta, err := svc.Take("late cancel")
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("Take = (%v, %v), want context.Canceled", meta, err)
	}
	entries, _ := os.ReadDir(filepath.Join(svc.s.ConfigDir(), "snapshots"))
	if len(entries) != 0 {
		t.Errorf("a cancelled snapshot must leave no file, found %d", len(entries))
	}
}

// Progress events carry the operation identity so the UI can route a cancel.
func TestSnapshotProgressCarriesOperation(t *testing.T) {
	var sawOp bool
	SetEventSink(func(name string, data map[string]any) {
		if name == "snapshot:progress" && data["opId"] != "" && data["opKind"] == "snapshot" {
			sawOp = true
		}
	})
	t.Cleanup(func() { SetEventSink(nil) })
	if _, err := snapshotHarness(t, defaultSnapState(), nil).Take("stamped"); err != nil {
		t.Fatal(err)
	}
	if !sawOp {
		t.Error("snapshot:progress events must be stamped with opId/opKind")
	}
}
