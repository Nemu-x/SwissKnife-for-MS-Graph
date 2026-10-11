// Package ops gives every long-running operation an identity (opId), its own
// cancellable context, and a single-flight guarantee per operation kind.
// Progress events stamp the opId so concurrent operations never mix in the UI,
// and Cancel aborts in-flight HTTP via context instead of waiting for the
// current item to finish (ADR-006 groundwork).
package ops

import (
	"context"
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"sync"
	"time"
)

// Kind labels a family of long-running operations. At most one operation of a
// kind is live at a time (single-flight) — the UI mirrors this with one primary
// job slot per kind.
type Kind string

const (
	KindTransfer Kind = "transfer"
	KindMailbox  Kind = "mailbox" // mailbox copy through the import/export APIs
	KindPlaybook Kind = "playbook"
	KindCleanup  Kind = "cleanup"
	KindMirror   Kind = "mirror"
	KindBulk     Kind = "bulk"
	KindUpdate   Kind = "update"
	KindSnapshot Kind = "snapshot" // tenant configuration snapshot
	KindDrift    Kind = "drift"    // background drift check against a baseline
	KindFanOut   Kind = "fanout"   // read action across several tenants
	KindAction   Kind = "action"   // catalog action apply (internal/engine)
	KindPreview  Kind = "preview"  // catalog action preview (Plan); several may run
)

// WriteKinds are the operations that change a tenant or run long against
// one; a connection switch is refused while one is live.
var WriteKinds = []Kind{KindTransfer, KindMailbox, KindPlaybook, KindCleanup, KindMirror, KindBulk, KindSnapshot, KindAction}

// AlreadyRunningError signals a single-flight violation the UI can explain.
type AlreadyRunningError struct{ Kind Kind }

func (e *AlreadyRunningError) Error() string {
	return fmt.Sprintf("a %s operation is already running — wait for it to finish or cancel it", e.Kind)
}

// Operation is one live run: identity plus a context that cancels with it.
type Operation struct {
	ID        string
	Kind      Kind
	Ctx       context.Context
	StartedAt time.Time
	cancel    context.CancelFunc
}

// Canceled reports whether the operation's context has been cancelled — the
// fast-path check services use between items.
func (o *Operation) Canceled() bool { return o.Ctx.Err() != nil }

// Registry tracks live operations. One registry per app session.
type Registry struct {
	mu   sync.Mutex
	live map[Kind]*Operation
	byID map[string]*Operation
}

func NewRegistry() *Registry {
	return &Registry{live: map[Kind]*Operation{}, byID: map[string]*Operation{}}
}

// Start mints an operation of the kind, deriving its context from parent
// (session context for top-level runs, an op context for child operations —
// cancelling the parent then cancels the child too). Returns
// *AlreadyRunningError when an operation of the kind is already live.
func (r *Registry) Start(parent context.Context, kind Kind) (*Operation, error) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if _, busy := r.live[kind]; busy {
		return nil, &AlreadyRunningError{Kind: kind}
	}
	ctx, cancel := context.WithCancel(parent)
	op := &Operation{ID: newID(), Kind: kind, Ctx: ctx, StartedAt: time.Now(), cancel: cancel}
	r.live[kind] = op
	r.byID[op.ID] = op
	return op, nil
}

// Track registers a cancellable operation that does not take its kind's
// single-flight slot (previews: several may run at once). id names it so the
// caller can cancel it before the call returns ("" mints one); an id that is
// already live is refused.
func (r *Registry) Track(parent context.Context, kind Kind, id string) (*Operation, error) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if id == "" {
		id = newID()
	}
	if _, dup := r.byID[id]; dup {
		return nil, fmt.Errorf("operation %s is already running", id)
	}
	ctx, cancel := context.WithCancel(parent)
	op := &Operation{ID: id, Kind: kind, Ctx: ctx, StartedAt: time.Now(), cancel: cancel}
	r.byID[id] = op
	return op, nil
}

// Busy returns a live operation of one of the kinds (single-flight ones
// only), or nil.
func (r *Registry) Busy(kinds ...Kind) *Operation {
	r.mu.Lock()
	defer r.mu.Unlock()
	for _, k := range kinds {
		if op := r.live[k]; op != nil {
			return op
		}
	}
	return nil
}

// Finish cancels the operation's context and frees its single-flight slot.
// Always call it (defer) when the run returns.
func (r *Registry) Finish(op *Operation) {
	if op == nil {
		return
	}
	r.mu.Lock()
	defer r.mu.Unlock()
	op.cancel()
	if cur := r.live[op.Kind]; cur == op {
		delete(r.live, op.Kind)
	}
	// A tracked id may be reused once finished: never drop a newer op.
	if r.byID[op.ID] == op {
		delete(r.byID, op.ID)
	}
}

// Cancel cancels a live operation by id; an unknown id is a no-op.
func (r *Registry) Cancel(id string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if op, ok := r.byID[id]; ok {
		op.cancel()
	}
}

// CancelKind cancels the live operation of the kind (legacy per-kind Cancel
// buttons) and every tracked one of the kind; a kind with no live operation
// is a no-op.
func (r *Registry) CancelKind(kind Kind) {
	r.mu.Lock()
	defer r.mu.Unlock()
	if op, ok := r.live[kind]; ok {
		op.cancel()
	}
	for _, op := range r.byID {
		if op.Kind == kind {
			op.cancel()
		}
	}
}

// Live returns the live operation of a kind, or nil.
func (r *Registry) Live(kind Kind) *Operation {
	r.mu.Lock()
	defer r.mu.Unlock()
	return r.live[kind]
}

// newID is time-sortable: unix-millis hex plus random suffix.
func newID() string {
	var b [6]byte
	_, _ = rand.Read(b[:])
	return fmt.Sprintf("%011x-%s", time.Now().UnixMilli(), hex.EncodeToString(b[:]))
}
