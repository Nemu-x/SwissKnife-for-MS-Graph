package auth

import (
	"crypto/x509"
	"testing"
	"time"
)

func TestGeneratedCertificateRoundTrips(t *testing.T) {
	g, err := GenerateCertificate("SwissKnifeGraph-test")
	if err != nil {
		t.Fatal(err)
	}
	if len(g.Thumbprint) != 40 || g.Password == "" {
		t.Fatalf("thumbprint %q password empty=%v", g.Thumbprint, g.Password == "")
	}
	if d := time.Until(g.NotAfter); d < 700*24*time.Hour {
		t.Fatalf("expires too soon: %v", g.NotAfter)
	}
	c, err := x509.ParseCertificate(g.CER)
	if err != nil || c.Subject.CommonName != "SwissKnifeGraph-test" {
		t.Fatalf("cer: %v %v", c, err)
	}
	// The PFX must open with its password and build a credential.
	if _, err := NewClientCertificate("tenant", "00000000-0000-0000-0000-000000000000", g.PFX, g.Password); err != nil {
		t.Fatalf("pfx: %v", err)
	}
	if _, err := NewClientCertificate("tenant", "client", g.PFX, "wrong"); err == nil {
		t.Fatal("a wrong password must fail")
	}
}
