package services

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/google/uuid"
	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

// On-premises Active Directory connections (ADR-008 M5). They are separate
// from tenant profiles: a hybrid admin works in both at once, and the AD
// actions need no tenant connection. Settings live in onprem.json, the bind
// password in the OS keychain.

// OnPremConnection is a saved directory connection.
type OnPremConnection struct {
	ldapx.Config
	HasPassword bool `json:"hasPassword"`
}

// OnPremStatus is the active directory connection, if any.
type OnPremStatus struct {
	Connected bool   `json:"connected"`
	ID        string `json:"id,omitempty"`
	Name      string `json:"name,omitempty"`
	Host      string `json:"host,omitempty"`
	Secure    bool   `json:"secure"`
}

// directories holds the active directory client per session.
var directories sync.Map // *session.Session → *ldapx.Client

func directoryFor(s *session.Session) *ldapx.Client {
	if v, ok := directories.Load(s); ok {
		return v.(*ldapx.Client)
	}
	return nil
}

var onpremMu sync.Mutex // onprem.json read-modify-write

func onpremKey(id string) string { return "onprem:" + id }

// OnPremService manages the connections and the active one.
type OnPremService struct{ s *session.Session }

func NewOnPremService(s *session.Session) *OnPremService { return &OnPremService{s: s} }

func (o *OnPremService) path() (string, error) {
	dir := o.s.ConfigDir()
	if dir == "" {
		return "", errors.New("config directory is not set")
	}
	return filepath.Join(dir, "onprem.json"), nil
}

func (o *OnPremService) load() ([]ldapx.Config, error) {
	p, err := o.path()
	if err != nil {
		return nil, err
	}
	b, err := os.ReadFile(p)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	var list []ldapx.Config
	return list, json.Unmarshal(b, &list)
}

func (o *OnPremService) save(list []ldapx.Config) error {
	p, err := o.path()
	if err != nil {
		return err
	}
	b, err := json.MarshalIndent(list, "", "  ")
	if err != nil {
		return err
	}
	tmp := p + ".tmp"
	if err := os.WriteFile(tmp, b, 0o600); err != nil {
		return err
	}
	return os.Rename(tmp, p)
}

// Connections lists the saved connections.
func (o *OnPremService) Connections() ([]OnPremConnection, error) {
	onpremMu.Lock()
	defer onpremMu.Unlock()
	list, err := o.load()
	out := []OnPremConnection{}
	for _, c := range list {
		_, has := secrets.NamedSecret(onpremKey(c.ID))
		out = append(out, OnPremConnection{Config: c, HasPassword: has})
	}
	return out, err
}

// SaveConnection creates or updates a connection; an empty password keeps
// the stored one.
func (o *OnPremService) SaveConnection(c ldapx.Config, password string) (*OnPremConnection, error) {
	c.Host, c.BaseDN, c.BindDN, c.Name = strings.TrimSpace(c.Host), strings.TrimSpace(c.BaseDN), strings.TrimSpace(c.BindDN), strings.TrimSpace(c.Name)
	if c.Name == "" {
		c.Name = c.Host
	}
	if err := c.Validate(); err != nil {
		return nil, err
	}
	onpremMu.Lock()
	defer onpremMu.Unlock()
	list, err := o.load()
	if err != nil {
		return nil, err
	}
	if c.ID == "" {
		c.ID = uuid.NewString()
	}
	prev := append([]ldapx.Config(nil), list...) // restored if the password cannot be stored
	replaced := false
	for i := range list {
		if list[i].ID == c.ID {
			list[i], replaced = c, true
		}
	}
	if !replaced {
		list = append(list, c)
	}
	// The settings are saved first, so a failed save leaves no password in
	// the keychain that no connection refers to.
	if err := o.save(list); err != nil {
		return nil, err
	}
	if password != "" {
		if err := secrets.SetNamedSecret(onpremKey(c.ID), password); err != nil {
			// Settings and password change together or not at all.
			if rerr := o.save(prev); rerr != nil {
				return nil, errors.Join(err, fmt.Errorf("restoring the previous settings: %w", rerr))
			}
			return nil, err
		}
	}
	// Edited while active: the next operation uses the new settings.
	if cur := directoryFor(o.s); cur != nil && cur.Config().ID == c.ID {
		directories.Delete(o.s)
	}
	_, has := secrets.NamedSecret(onpremKey(c.ID))
	return &OnPremConnection{Config: c, HasPassword: has}, nil
}

// DeleteConnection removes a connection and its password.
func (o *OnPremService) DeleteConnection(id string) error {
	onpremMu.Lock()
	defer onpremMu.Unlock()
	list, err := o.load()
	if err != nil {
		return err
	}
	out := list[:0]
	for _, c := range list {
		if c.ID != id {
			out = append(out, c)
		}
	}
	if err := o.save(out); err != nil {
		return err
	}
	if cur := directoryFor(o.s); cur != nil && cur.Config().ID == id {
		directories.Delete(o.s)
	}
	return secrets.SetNamedSecret(onpremKey(id), "")
}

func (o *OnPremService) client(id string) (*ldapx.Client, error) {
	onpremMu.Lock()
	list, err := o.load()
	onpremMu.Unlock()
	if err != nil {
		return nil, err
	}
	for _, c := range list {
		if c.ID == id {
			pw, ok := secrets.NamedSecret(onpremKey(id))
			if !ok {
				return nil, errors.New("no bind password saved for this connection")
			}
			return ldapx.New(c, pw), nil
		}
	}
	return nil, errors.New("connection not found")
}

// Connect tests a connection (bind) and makes it the active one.
func (o *OnPremService) Connect(id string) (*OnPremStatus, error) {
	c, err := o.client(id)
	if err != nil {
		return nil, err
	}
	ctx, cancel := context.WithTimeout(o.s.Ctx(), 30*time.Second)
	defer cancel()
	err = c.Test(ctx)
	o.s.Record("onprem.connect", c.Config().Host, "name="+c.Config().Name+" tls="+string(c.Config().TLS), err)
	if err != nil {
		return nil, err
	}
	directories.Store(o.s, c)
	return o.Status(), nil
}

// Disconnect forgets the active connection.
func (o *OnPremService) Disconnect() {
	if c := directoryFor(o.s); c != nil {
		o.s.Record("onprem.disconnect", c.Config().Host, "", nil)
	}
	directories.Delete(o.s)
}

// Status reports the active connection.
func (o *OnPremService) Status() *OnPremStatus {
	c := directoryFor(o.s)
	if c == nil {
		return &OnPremStatus{}
	}
	cfg := c.Config()
	return &OnPremStatus{Connected: true, ID: cfg.ID, Name: cfg.Name, Host: cfg.Host, Secure: c.Secure()}
}

// PickCAFile chooses the domain CA certificate (PEM).
func (o *OnPremService) PickCAFile() (string, error) {
	if o.s.Ctx().Value("events") == nil {
		return "", errors.New("file dialogs need the desktop app")
	}
	return wrt.OpenFileDialog(o.s.Ctx(), wrt.OpenDialogOptions{
		Title:   "Domain CA certificate",
		Filters: []wrt.FileFilter{{DisplayName: "Certificate (*.pem, *.crt, *.cer)", Pattern: "*.pem;*.crt;*.cer"}},
	})
}
