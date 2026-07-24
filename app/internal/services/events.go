package services

import (
	"context"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"
)

// emitEvent forwards to the Wails event bus. Headless runs (unit tests, CLI)
// carry a bare context without the Wails frontend attached; the runtime would
// log.Fatal in that case, so we only emit when the event bus is present.
func emitEvent(ctx context.Context, name string, data any) {
	if ctx == nil || ctx.Value("events") == nil {
		return
	}
	wrt.EventsEmit(ctx, name, data)
}
