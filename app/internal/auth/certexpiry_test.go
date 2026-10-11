package auth

import (
	"crypto/rand"
	"crypto/rsa"
	"crypto/x509"
	"crypto/x509/pkix"
	"math/big"
	"strings"
	"testing"
	"time"

	pkcs12 "software.sslmate.com/src/go-pkcs12"
)

// pfxValidUntil builds a password-protected PFX whose certificate ends at notAfter.
func pfxValidUntil(t *testing.T, notAfter time.Time) []byte {
	t.Helper()
	key, err := rsa.GenerateKey(rand.Reader, 2048)
	if err != nil {
		t.Fatal(err)
	}
	tmpl := &x509.Certificate{
		SerialNumber: big.NewInt(1),
		Subject:      pkix.Name{CommonName: "expiry-test"},
		NotBefore:    notAfter.AddDate(-1, 0, 0),
		NotAfter:     notAfter,
	}
	der, err := x509.CreateCertificate(rand.Reader, tmpl, tmpl, &key.PublicKey, key)
	if err != nil {
		t.Fatal(err)
	}
	cert, _ := x509.ParseCertificate(der)
	pfx, err := pkcs12.Modern.Encode(key, cert, nil, "pw")
	if err != nil {
		t.Fatal(err)
	}
	return pfx
}

func TestExpiredCertificateIsRefusedWithAHint(t *testing.T) {
	expired := time.Now().AddDate(0, 0, -3).Truncate(time.Second)
	_, err := NewClientCertificate("tenant", "client", pfxValidUntil(t, expired), "pw")
	if !IsCertExpired(err) {
		t.Fatalf("err = %v, want a CertExpiredError", err)
	}
	if msg := err.Error(); !strings.Contains(msg, "expired on "+expired.Format("2006-01-02")) || !strings.Contains(msg, "Generate") {
		t.Errorf("message does not say when and how to regenerate: %q", msg)
	}
	// Expiring soon still connects.
	if _, err := NewClientCertificate("tenant", "client", pfxValidUntil(t, time.Now().AddDate(0, 0, 5)), "pw"); err != nil {
		t.Errorf("valid certificate refused: %v", err)
	}
}

func TestCertificateNotAfter(t *testing.T) {
	want := time.Now().AddDate(0, 0, 20).UTC().Truncate(time.Second)
	got, err := CertificateNotAfter(pfxValidUntil(t, want), "pw")
	if err != nil || !got.Equal(want) {
		t.Fatalf("NotAfter = %v, %v; want %v", got, err, want)
	}
	if _, err := CertificateNotAfter(pfxValidUntil(t, want), "wrong"); err == nil {
		t.Error("a wrong password must fail")
	}
}
