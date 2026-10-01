package services

import (
	"context"
	"sync"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/ops"
)

// eventSink, when non-nil, receives events instead of the Wails bus. Tests use
// it to capture emissions; production leaves it nil.
var (
	eventSinkMu sync.RWMutex
	eventSink   func(name string, data map[string]any)
)

// SetEventSink installs a process-wide receiver for map-shaped events, used by
// headless front ends (the CLI) to stream progress that the GUI gets over the
// Wails bus. nil restores the default (Wails bus when attached, else dropped).
func SetEventSink(fn func(name string, data map[string]any)) {
	eventSinkMu.Lock()
	eventSink = fn
	eventSinkMu.Unlock()
}

// emitEvent forwards to the Wails event bus. Headless runs (unit tests, CLI)
// carry a bare context without the Wails frontend attached; the runtime would
// log.Fatal in that case, so we only emit when the event bus is present.
func emitEvent(ctx context.Context, name string, data any) {
	eventSinkMu.RLock()
	sink := eventSink
	eventSinkMu.RUnlock()
	if m, ok := data.(map[string]any); ok && sink != nil {
		sink(name, m)
		return
	}
	if ctx == nil || ctx.Value("events") == nil {
		return
	}
	wrt.EventsEmit(ctx, name, data)
}

// emitOp stamps the operation identity into the payload and emits it. Every
// operation progress event MUST go through here so the UI can demultiplex
// concurrent operations by opId. A nil op degrades to an unstamped emit on the
// given fallback context (quick one-shot actions without a registered op).
func emitOp(fallback context.Context, op *ops.Operation, name string, data map[string]any) {
	if data == nil {
		data = map[string]any{}
	}
	ctx := fallback
	if op != nil {
		data["opId"] = op.ID
		data["opKind"] = string(op.Kind)
		ctx = op.Ctx
	}
	emitEvent(ctx, name, data)
}
