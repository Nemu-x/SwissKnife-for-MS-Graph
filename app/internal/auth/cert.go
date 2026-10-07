package auth

import (
	"crypto"
	"crypto/rand"
	"crypto/rsa"
	"crypto/sha1" // the certificate thumbprint Entra shows is SHA-1 by definition
	"crypto/x509"
	"crypto/x509/pkix"
	"encoding/base64"
	"encoding/hex"
	"fmt"
	"math/big"
	"strings"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	pkcs12 "software.sslmate.com/src/go-pkcs12"
)

// NewClientCertificate authenticates app-only with a certificate: the PFX
// bytes and their password (empty for an unprotected PFX). Exchange Online
// PowerShell accepts only certificates for app-only sign-in (ADR-008).
func NewClientCertificate(tenantID, clientID string, pfx []byte, password string) (*TokenProvider, error) {
	certs, key, err := parseCertificate(pfx, password)
	if err != nil {
		return nil, fmt.Errorf("read certificate: %w", err)
	}
	cred, err := azidentity.NewClientCertificateCredential(tenantID, clientID, certs, key, nil)
	if err != nil {
		return nil, err
	}
	return &TokenProvider{cred: cred, tokens: map[string]azcore.AccessToken{}}, nil
}

// parseCertificate reads a PFX in both the modern (AES/SHA-256) and the legacy
// encoding — azidentity's own parser knows only the legacy one — and falls
// back to azidentity for PEM files.
func parseCertificate(data []byte, password string) ([]*x509.Certificate, crypto.PrivateKey, error) {
	key, cert, chain, err := pkcs12.DecodeChain(data, password)
	if err == nil {
		return append([]*x509.Certificate{cert}, chain...), key, nil
	}
	if certs, pemKey, perr := azidentity.ParseCertificates(data, []byte(password)); perr == nil {
		return certs, pemKey, nil
	}
	return nil, nil, err
}

// GeneratedCert is a fresh self-signed certificate for an app registration.
type GeneratedCert struct {
	PFX        []byte    // certificate + private key, protected by Password
	CER        []byte    // public certificate (DER) to upload to Entra
	Password   string    // random; stored in the OS keychain with the profile
	Thumbprint string    // SHA-1, upper-case hex — what Entra lists
	NotAfter   time.Time // expiry
}

// GenerateCertificate creates an RSA-2048 self-signed certificate valid for
// two years, the shape Entra app registrations accept for client auth.
func GenerateCertificate(subject string) (*GeneratedCert, error) {
	key, err := rsa.GenerateKey(rand.Reader, 2048)
	if err != nil {
		return nil, err
	}
	serial, err := rand.Int(rand.Reader, new(big.Int).Lsh(big.NewInt(1), 127))
	if err != nil {
		return nil, err
	}
	now := time.Now().Add(-5 * time.Minute) // tolerate clock skew
	tmpl := &x509.Certificate{
		SerialNumber: serial,
		Subject:      pkix.Name{CommonName: subject},
		NotBefore:    now,
		NotAfter:     now.AddDate(2, 0, 0),
		KeyUsage:     x509.KeyUsageDigitalSignature | x509.KeyUsageKeyEncipherment,
		ExtKeyUsage:  []x509.ExtKeyUsage{x509.ExtKeyUsageClientAuth},
	}
	der, err := x509.CreateCertificate(rand.Reader, tmpl, tmpl, &key.PublicKey, key)
	if err != nil {
		return nil, err
	}
	cert, err := x509.ParseCertificate(der)
	if err != nil {
		return nil, err
	}
	pw := make([]byte, 24)
	if _, err := rand.Read(pw); err != nil {
		return nil, err
	}
	password := base64.RawURLEncoding.EncodeToString(pw)
	pfx, err := pkcs12.Modern.Encode(key, cert, nil, password)
	if err != nil {
		return nil, err
	}
	sum := sha1.Sum(der) // thumbprint, not a security decision
	return &GeneratedCert{
		PFX:        pfx,
		CER:        der,
		Password:   password,
		Thumbprint: strings.ToUpper(hex.EncodeToString(sum[:])),
		NotAfter:   cert.NotAfter,
	}, nil
}
