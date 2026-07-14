// Package secrets stores connection profiles. Non-secret data (tenant/client id, mode)
// lives in profiles.json under AppData; the client secret lives only in the OS keychain (ADR-002).
package secrets

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"

	"github.com/google/uuid"
	"github.com/zalando/go-keyring"
)

const keyringService = "SwissKnifeGraph"

// Profile is a saved connection (without the secret).
type Profile struct {
	ID       string `json:"id"`
	Name     string `json:"name"`
	TenantID string `json:"tenantId"`
	ClientID string `json:"clientId"`
	AuthMode string `json:"authMode"`  // client_secret | device_code
	HasSecret bool  `json:"hasSecret"` // whether a secret exists in the keychain
}

type Store struct {
	dir  string
	path string
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
	return os.WriteFile(s.path, data, 0o600)
}

func (s *Store) List() ([]Profile, error) {
	return s.load()
}

// Save creates/updates a profile. An empty secret means keep the stored one.
func (s *Store) Save(p Profile, secret string) (Profile, error) {
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

// Secret returns the profile secret from the keychain.
func (s *Store) Secret(profileID string) (string, error) {
	v, err := keyring.Get(keyringService, profileID)
	if errors.Is(err, keyring.ErrNotFound) {
		return "", errors.New("secret not found in keychain — re-enter it in profile settings")
	}
	return v, err
}

func (s *Store) Delete(profileID string) error {
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
