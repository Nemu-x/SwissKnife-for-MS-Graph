package main

import (
	"context"

	"swissknife-app/internal/services"
	"swissknife-app/internal/session"
)

// version is injected at build time via -ldflags "-X main.version=...".
var version = "dev"

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
	return version
}
