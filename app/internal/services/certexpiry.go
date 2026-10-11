package services

import (
	"os"
	"sync"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/secrets"
)

// certExpiryKey identifies one version of a PFX file: a replaced file (new
// size or modification time) is read again.
type certExpiryKey struct {
	path string
	size int64
	mod  time.Time
}

// certExpiryCache spares a keychain read and a PFX decode per profile on every
// profile listing.
var certExpiryCache sync.Map // certExpiryKey -> time.Time

// annotateCertExpiry fills CertNotAfter on certificate profiles, so the UI can
// warn before the certificate expires. Best-effort: a missing file or a
// password the keychain no longer has leaves the field empty (connecting
// reports those problems with their own message).
func annotateCertExpiry(store *secrets.Store, list []secrets.Profile) {
	for i := range list {
		p := &list[i]
		if p.AuthMode != string(auth.ModeClientCertificate) || p.CertPath == "" || !p.HasSecret {
			continue
		}
		st, err := os.Stat(p.CertPath)
		if err != nil {
			continue
		}
		key := certExpiryKey{p.CertPath, st.Size(), st.ModTime()}
		if v, ok := certExpiryCache.Load(key); ok {
			na := v.(time.Time)
			p.CertNotAfter = &na
			continue
		}
		password, err := store.Secret(p.ID)
		if err != nil {
			continue
		}
		data, err := os.ReadFile(p.CertPath)
		if err != nil {
			continue
		}
		na, err := auth.CertificateNotAfter(data, password)
		if err != nil {
			continue
		}
		certExpiryCache.Store(key, na)
		p.CertNotAfter = &na
	}
}
