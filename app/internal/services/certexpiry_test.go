package services

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/zalando/go-keyring"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

func TestProfilesCarryCertificateExpiry(t *testing.T) {
	keyring.MockInit()
	dir := t.TempDir()
	store := secrets.NewStoreAt(dir)
	c := NewConnectService(session.New(auditlog.New(dir)), store)

	g, err := auth.GenerateCertificate("expiry-test")
	if err != nil {
		t.Fatal(err)
	}
	pfx := filepath.Join(dir, "app.pfx")
	if err := os.WriteFile(pfx, g.PFX, 0o600); err != nil {
		t.Fatal(err)
	}
	if _, err := store.Save(secrets.Profile{Name: "cert", TenantID: "t", ClientID: "c",
		AuthMode: string(auth.ModeClientCertificate), CertPath: pfx}, g.Password); err != nil {
		t.Fatal(err)
	}
	if _, err := store.Save(secrets.Profile{Name: "secret", TenantID: "t", ClientID: "c",
		AuthMode: string(auth.ModeClientSecret)}, "s3cret"); err != nil {
		t.Fatal(err)
	}

	list, err := c.Profiles()
	if err != nil {
		t.Fatal(err)
	}
	for _, p := range list {
		switch p.Name {
		case "cert":
			if p.CertNotAfter == nil || !p.CertNotAfter.Equal(g.NotAfter) {
				t.Errorf("certNotAfter = %v, want %v", p.CertNotAfter, g.NotAfter)
			}
			// Saving a listed profile back must not persist the derived field.
			if _, err := store.Save(p, ""); err != nil {
				t.Fatal(err)
			}
		case "secret":
			if p.CertNotAfter != nil {
				t.Errorf("secret profile got an expiry: %v", p.CertNotAfter)
			}
		}
	}
	raw, _ := os.ReadFile(filepath.Join(dir, "profiles.json"))
	if strings.Contains(string(raw), "certNotAfter") {
		t.Errorf("certNotAfter persisted: %s", raw)
	}

	// A replaced file is read again (the cache keys on size + mtime).
	if err := os.WriteFile(pfx, []byte("not a pfx"), 0o600); err != nil {
		t.Fatal(err)
	}
	_ = os.Chtimes(pfx, time.Now().Add(time.Hour), time.Now().Add(time.Hour))
	list, _ = c.Profiles()
	for _, p := range list {
		if p.Name == "cert" && p.CertNotAfter != nil {
			t.Errorf("unreadable file still reports an expiry: %v", p.CertNotAfter)
		}
	}
}
