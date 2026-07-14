package main

import (
	"context"

	"swissknife-app/internal/session"
)

// version is injected at build time via -ldflags "-X main.version=...".
var version = "dev"

// App is the root binding: lifecycle and metadata.
type App struct {
	session *session.Session
}

func NewApp(s *session.Session) *App {
	return &App{session: s}
}

func (a *App) startup(ctx context.Context) {
	a.session.SetAppContext(ctx)
}

func (a *App) Version() string {
	return version
}
