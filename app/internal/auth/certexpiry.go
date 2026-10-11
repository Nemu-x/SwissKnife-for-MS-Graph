package auth

import (
	"errors"
	"fmt"
	"time"
)

// CertExpiringWindow is how early the UI starts warning about an expiring
// client certificate.
const CertExpiringWindow = 30 * 24 * time.Hour

// CertExpiredError means the client certificate is past its NotAfter: Entra
// rejects it, so connecting would only fail later with an opaque AADSTS error.
type CertExpiredError struct {
	NotAfter time.Time
}

func (e *CertExpiredError) Error() string {
	return fmt.Sprintf("the client certificate expired on %s — generate a new certificate in the connection form "+
		"(Certificate → Generate), upload its .cer to the app registration in Entra (Certificates & secrets) "+
		"and save the profile with the new .pfx", e.NotAfter.Format("2006-01-02"))
}

// IsCertExpired reports whether err is (or wraps) a CertExpiredError.
func IsCertExpired(err error) bool {
	var ce *CertExpiredError
	return errors.As(err, &ce)
}

// CertificateNotAfter returns the expiry of the leaf certificate in a PFX (or
// PEM) file, which needs the file's password to read.
func CertificateNotAfter(data []byte, password string) (time.Time, error) {
	certs, _, err := parseCertificate(data, password)
	if err != nil {
		return time.Time{}, err
	}
	if len(certs) == 0 || certs[0] == nil {
		return time.Time{}, errors.New("no certificate in the file")
	}
	return certs[0].NotAfter, nil
}

// checkNotExpired refuses a certificate whose validity ended before now.
func checkNotExpired(notAfter, now time.Time) error {
	if !notAfter.IsZero() && now.After(notAfter) {
		return &CertExpiredError{NotAfter: notAfter}
	}
	return nil
}
