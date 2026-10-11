package services

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

// RevealCertificate hands its path to a shell command, so anything outside
// the app's certs folder must be refused before that.
func TestRevealCertificateRefusesForeignPaths(t *testing.T) {
	dir := t.TempDir()
	c := NewConnectService(session.New(auditlog.New(dir)), secrets.NewStoreAt(dir))
	outside := filepath.Join(dir, "profiles.json")
	_ = os.WriteFile(outside, []byte("[]"), 0o600)
	for _, p := range []string{
		outside,
		filepath.Join(dir, "certs", "..", "profiles.json"),
		filepath.Join(dir, "certs"),
		"-aCalculator",
		"",
	} {
		if err := c.RevealCertificate(p); err == nil || !strings.Contains(err.Error(), "not a certificate") {
			t.Errorf("RevealCertificate(%q) = %v, want refusal", p, err)
		}
	}
}
