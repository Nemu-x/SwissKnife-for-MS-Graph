// Package worker runs PowerShell for another SwissKnife over the network
// (ADR-008 D9): a Windows machine with the modules installed executes the
// allow-listed cmdlet calls of a paired client — e.g. a Mac without
// PowerShell, or Security & Compliance cmdlets that only run on Windows.
//
// Trust is pairwise and pinned: each side has a self-signed certificate; a
// one-time code shown by the worker binds the two during pairing, after
// which every connection is mutual TLS against the pinned fingerprints.
// Tokens travel with each call and are never written to disk.
package worker

import (
	"crypto/ecdsa"
	"crypto/elliptic"
	"crypto/rand"
	"crypto/sha256"
	"crypto/tls"
	"crypto/x509"
	"crypto/x509/pkix"
	"encoding/hex"
	"encoding/pem"
	"errors"
	"math/big"
	"os"
	"path/filepath"
	"time"
)

// Identity is one side's certificate and its fingerprint.
type Identity struct {
	Cert        tls.Certificate
	Fingerprint string // SHA-256 of the DER certificate, hex
}

// Fingerprint of a DER certificate.
func Fingerprint(der []byte) string {
	sum := sha256.Sum256(der)
	return hex.EncodeToString(sum[:])
}

// LoadOrCreateIdentity reads <dir>/<name>.crt/.key or creates them (P-256,
// ten years; the key file is readable by the user only).
func LoadOrCreateIdentity(dir, name string) (*Identity, error) {
	certPath, keyPath := filepath.Join(dir, name+".crt"), filepath.Join(dir, name+".key")
	if c, err := tls.LoadX509KeyPair(certPath, keyPath); err == nil {
		return &Identity{Cert: c, Fingerprint: Fingerprint(c.Certificate[0])}, nil
	} else if !errors.Is(err, os.ErrNotExist) {
		return nil, err
	}
	key, err := ecdsa.GenerateKey(elliptic.P256(), rand.Reader)
	if err != nil {
		return nil, err
	}
	serial, _ := rand.Int(rand.Reader, new(big.Int).Lsh(big.NewInt(1), 62))
	host, _ := os.Hostname()
	tmpl := &x509.Certificate{
		SerialNumber: serial, Subject: pkix.Name{CommonName: "SwissKnife " + name + " " + host},
		NotBefore: time.Now().Add(-time.Hour), NotAfter: time.Now().AddDate(10, 0, 0),
		KeyUsage:    x509.KeyUsageDigitalSignature,
		ExtKeyUsage: []x509.ExtKeyUsage{x509.ExtKeyUsageServerAuth, x509.ExtKeyUsageClientAuth},
	}
	der, err := x509.CreateCertificate(rand.Reader, tmpl, tmpl, &key.PublicKey, key)
	if err != nil {
		return nil, err
	}
	kder, err := x509.MarshalECPrivateKey(key)
	if err != nil {
		return nil, err
	}
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return nil, err
	}
	if err := os.WriteFile(keyPath, pem.EncodeToMemory(&pem.Block{Type: "EC PRIVATE KEY", Bytes: kder}), 0o600); err != nil {
		return nil, err
	}
	if err := os.WriteFile(certPath, pem.EncodeToMemory(&pem.Block{Type: "CERTIFICATE", Bytes: der}), 0o644); err != nil {
		return nil, err
	}
	c, err := tls.LoadX509KeyPair(certPath, keyPath)
	if err != nil {
		return nil, err
	}
	return &Identity{Cert: c, Fingerprint: Fingerprint(der)}, nil
}

// pinned verifies that the peer presented exactly one of the pinned
// certificates (no chain, no CA: the fingerprint is the trust).
func pinned(allowed func(fp string) bool) func([][]byte, [][]*x509.Certificate) error {
	return func(raw [][]byte, _ [][]*x509.Certificate) error {
		if len(raw) == 0 {
			return errors.New("no certificate")
		}
		if !allowed(Fingerprint(raw[0])) {
			return errors.New("certificate not paired")
		}
		return nil
	}
}
