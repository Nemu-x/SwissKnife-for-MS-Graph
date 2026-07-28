package ops

import (
	"context"
	"errors"
	"testing"
	"time"
)

func TestSingleFlightPerKind(t *testing.T) {
	r := NewRegistry()
	op1, err := r.Start(context.Background(), KindTransfer)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := r.Start(context.Background(), KindTransfer); err == nil {
		t.Fatal("second transfer op must be rejected")
	} else {
		var are *AlreadyRunningError
		if !errors.As(err, &are) || are.Kind != KindTransfer {
			t.Fatalf("want AlreadyRunningError{transfer}, got %v", err)
		}
	}
	// A different kind is fine while transfer is live.
	if _, err := r.Start(context.Background(), KindPlaybook); err != nil {
		t.Fatalf("other kinds must not be blocked: %v", err)
	}
	r.Finish(op1)
	if _, err := r.Start(context.Background(), KindTransfer); err != nil {
		t.Fatalf("slot must be free after Finish: %v", err)
	}
}

func TestCancelByIDAndKind(t *testing.T) {
	r := NewRegistry()
	op, _ := r.Start(context.Background(), KindPlaybook)
	if op.Canceled() {
		t.Fatal("fresh op must not be canceled")
	}
	r.Cancel("nope") // unknown id: no-op
	if op.Canceled() {
		t.Fatal("unknown id must not cancel anything")
	}
	r.Cancel(op.ID)
	if !op.Canceled() {
		t.Fatal("Cancel(id) must cancel the op context")
	}
	r.Finish(op)

	op2, _ := r.Start(context.Background(), KindPlaybook)
	r.CancelKind(KindPlaybook)
	if !op2.Canceled() {
		t.Fatal("CancelKind must cancel the live op")
	}
	r.Finish(op2)
}

func TestChildOpCancelsWithParent(t *testing.T) {
	r := NewRegistry()
	parent, _ := r.Start(context.Background(), KindPlaybook)
	child, err := r.Start(parent.Ctx, KindTransfer)
	if err != nil {
		t.Fatal(err)
	}
	r.CancelKind(KindPlaybook)
	if !child.Canceled() {
		t.Fatal("cancelling the parent playbook must cancel the child transfer")
	}
	r.Finish(child)
	r.Finish(parent)
}

func TestFinishFreesSlotOnlyForSameOp(t *testing.T) {
	r := NewRegistry()
	op1, _ := r.Start(context.Background(), KindCleanup)
	r.Finish(op1)
	op2, _ := r.Start(context.Background(), KindCleanup)
	r.Finish(op1) // stale finish must not free op2's slot
	if _, err := r.Start(context.Background(), KindCleanup); err == nil {
		t.Fatal("op2 must still hold the slot after a stale Finish")
	}
	r.Finish(op2)
}

func TestIDsAreSortable(t *testing.T) {
	a := newID()
	time.Sleep(2 * time.Millisecond) // force a later millisecond prefix
	b := newID()
	if a >= b {
		t.Fatalf("ids must sort by time: %s then %s", a, b)
	}
}
