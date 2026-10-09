package services

import (
	"os"
	"testing"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/session"
)

func TestMain(m *testing.M) {
	// Tests fake PowerShell through providers; pack scripts' own local check
	// would spawn the real one.
	packLocalPS = func(*session.Session, string) *engine.Reason { return nil }
	os.Exit(m.Run())
}
