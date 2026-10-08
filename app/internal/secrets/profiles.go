// Package secrets stores connection profiles. Non-secret data (tenant/client id, mode)
// lives in profiles.json under AppData; the client secret lives only in the OS keychain (ADR-002).
package secrets

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"

	"github.com/google/uuid"
	"github.com/zalando/go-keyring"

	"swissknife-app/internal/session"
)

const keyringService = "SwissKnifeGraph"

// Profile is a saved connection (without the secret).
type Profile struct {
	ID       string `json:"id"`
	Name     string `json:"name"`
	TenantID string `json:"tenantId"`
	ClientID string `json:"clientId"`
	AuthMode string `json:"authMode"`  // client_secret | device_code | client_certificate
	// CertPath is the PFX file of a client_certificate profile; its password
	// lives in the keychain in place of the client secret.
	CertPath string `json:"certPath,omitempty"`
	HasSecret bool  `json:"hasSecret"` // whether a secret exists in the keychain
	// DelegatedOrg is the customer tenant (id or domain) a partner signs in
	// to through GDAP; delegated (device code) profiles only.
	DelegatedOrg string `json:"delegatedOrg,omitempty"`
	// Policy limits what this profile may do once connected.
	Policy *session.Policy `json:"policy,omitempty"`
}

type Store struct {
	dir  string
	path string
	mu   sync.Mutex // serializes read-modify-write of profiles.json
}

func NewStore() (*Store, error) {
	base, err := os.UserConfigDir()
	if err != nil {
		return nil, err
	}
	dir := filepath.Join(base, "SwissKnifeGraph")
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return nil, err
	}
	return &Store{dir: dir, path: filepath.Join(dir, "profiles.json")}, nil
}

// NewStoreAt opens a store rooted at dir (tests, portable setups).
func NewStoreAt(dir string) *Store {
	return &Store{dir: dir, path: filepath.Join(dir, "profiles.json")}
}

// Dir is the app data directory (also used by the audit log).
func (s *Store) Dir() string { return s.dir }

func (s *Store) load() ([]Profile, error) {
	data, err := os.ReadFile(s.path)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	var out []Profile
	if err := json.Unmarshal(data, &out); err != nil {
		return nil, fmt.Errorf("profiles.json corrupted: %w", err)
	}
	return out, nil
}

func (s *Store) save(list []Profile) error {
	sort.Slice(list, func(i, j int) bool { return list[i].Name < list[j].Name })
	data, err := json.MarshalIndent(list, "", "  ")
	if err != nil {
		return err
	}
	tmp := s.path + ".tmp"
	if err := os.WriteFile(tmp, data, 0o600); err != nil {
		return err
	}
	return os.Rename(tmp, s.path)
}

func (s *Store) List() ([]Profile, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.load()
}

// Save creates/updates a profile. An empty secret means keep the stored one.
func (s *Store) Save(p Profile, secret string) (Profile, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	list, err := s.load()
	if err != nil {
		return Profile{}, err
	}

	if p.ID == "" {
		p.ID = uuid.NewString()
	}

	if secret != "" {
		if err := keyring.Set(keyringService, p.ID, secret); err != nil {
			return Profile{}, fmt.Errorf("keychain: %w", err)
		}
		p.HasSecret = true
	}

	replaced := false
	for i := range list {
		if list[i].ID == p.ID {
			if secret == "" {
				p.HasSecret = list[i].HasSecret
			}
			list[i] = p
			replaced = true
			break
		}
	}
	if !replaced {
		list = append(list, p)
	}
	return p, s.save(list)
}

// pendingKey names the keychain entry of a generated certificate's password
// before a profile owns it.
// The key hashes the absolute path: two folders may hold equally named files.
func pendingKey(certPath string) string {
	abs, err := filepath.Abs(certPath)
	if err != nil {
		abs = certPath
	}
	sum := sha256.Sum256([]byte(strings.ToLower(filepath.Clean(abs))))
	return "pending-cert:" + hex.EncodeToString(sum[:12])
}

// SetPendingCertPassword parks a generated PFX password in the keychain until
// a profile using that file is saved (the password never reaches the UI).
func (s *Store) SetPendingCertPassword(certPath, password string) error {
	return keyring.Set(keyringService, pendingKey(certPath), password)
}

// PendingCertPassword returns a parked password for the PFX file, if any.
func (s *Store) PendingCertPassword(certPath string) (string, bool) {
	if certPath == "" {
		return "", false
	}
	v, err := keyring.Get(keyringService, pendingKey(certPath))
	return v, err == nil
}

// DropPendingCertPassword removes a parked password once a profile owns it.
func (s *Store) DropPendingCertPassword(certPath string) {
	_ = keyring.Delete(keyringService, pendingKey(certPath))
}

// ClearSecret forgets a profile's stored secret (an auth-mode switch makes
// the old one meaningless: a client secret is not a PFX password).
func (s *Store) ClearSecret(profileID string) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	list, err := s.load()
	if err != nil {
		return err
	}
	if err := keyring.Delete(keyringService, profileID); err != nil && !errors.Is(err, keyring.ErrNotFound) {
		return err
	}
	for i := range list {
		if list[i].ID == profileID {
			list[i].HasSecret = false
		}
	}
	return s.save(list)
}

// Secret returns the profile secret from the keychain.
func (s *Store) Secret(profileID string) (string, error) {
	v, err := keyring.Get(keyringService, profileID)
	if errors.Is(err, keyring.ErrNotFound) {
		return "", errors.New("secret not found in keychain — re-enter it in profile settings")
	}
	return v, err
}

func (s *Store) Delete(profileID string) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	list, err := s.load()
	if err != nil {
		return err
	}
	out := list[:0]
	for _, p := range list {
		if p.ID != profileID {
			out = append(out, p)
		}
	}
	// remove the keychain secret regardless; ErrNotFound is not an error
	if err := keyring.Delete(keyringService, profileID); err != nil && !errors.Is(err, keyring.ErrNotFound) {
		return err
	}
	return s.save(out)
}

// Named keychain entries hold secrets that belong to no connection profile
// (an on-prem directory's bind password). An empty value deletes the entry.

// SetNamedSecret stores a secret under name.
func SetNamedSecret(name, value string) error {
	if value == "" {
		err := keyring.Delete(keyringService, name)
		if errors.Is(err, keyring.ErrNotFound) {
			return nil
		}
		return err
	}
	return keyring.Set(keyringService, name, value)
}

// NamedSecret returns a stored secret ("" and false when there is none).
func NamedSecret(name string) (string, bool) {
	v, err := keyring.Get(keyringService, name)
	return v, err == nil
}
