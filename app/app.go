package main

import (
	"context"

	"swissknife-app/internal/services"
	"swissknife-app/internal/session"
)

// baseVersion is the current release line. Bump it together with tagging a
// release — the Release workflow fails if the tag and this constant disagree.
const baseVersion = "0.9.0"

// version is injected at build time via -ldflags "-X main.version=...".
// Local builds (no ldflags) fall back to baseVersion with a -dev suffix
// (Wails builds with -buildvcs=false, so git metadata is not available).
var version = "dev"

func resolveVersion() string {
	if version != "dev" {
		return version
	}
	return baseVersion + "-dev"
}

// App is the root binding: lifecycle and metadata.
type App struct {
	session *session.Session
	update  *services.UpdateService
}

func NewApp(s *session.Session, update *services.UpdateService) *App {
	return &App{session: s, update: update}
}

func (a *App) startup(ctx context.Context) {
	a.session.SetAppContext(ctx)
	a.update.SetAppContext(ctx)
}

func (a *App) Version() string {
	return resolveVersion()
}
