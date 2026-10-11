// Package session holds the current connection and safety guards (ADR-002):
// read-only mode, typed confirm for destructive operations, audit-log recording.
package session

import (
	"context"
	"errors"
	"sync"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/ops"
)

// TokenBroker issues access tokens for resources other than Graph (Exchange,
// Teams) on behalf of the connected profile (ADR-008).
type TokenBroker interface {
	TokenFor(ctx context.Context, resource string) (string, error)
}

var ErrNotConnected = errors.New("not connected — connect to a tenant first")
var ErrReadOnly = errors.New("read-only mode is enabled — writes are blocked")

type Session struct {
	connMu       sync.Mutex // serialises Connect and Disconnect (hooks included)
	mu           sync.RWMutex
	ctx          context.Context
	client       *graphapi.Client
	tokens       TokenBroker
	tenantID     string
	appOnly      bool
	onDisconnect []func(prev *graphapi.Client)
	profileName  string
	readOnly     bool
	policy       Policy
	profileID    string
	delegatedOrg string
	scope        ScopeFunc
	configDir    string

	Audit   *auditlog.Log
	Ops     *ops.Registry
	Journal *journal.Log
}

func New(audit *auditlog.Log) *Session {
	return &Session{ctx: context.Background(), Audit: audit, Ops: ops.NewRegistry()}
}

// SetJournal attaches the persistent run journal (nil-safe: services check).
func (s *Session) SetJournal(j *journal.Log) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.Journal = j
}

// SetConfigDir stores the per-user config directory (%APPDATA%\SwissKnifeGraph)
// for services that persist small settings files (e.g. notifications).
func (s *Session) SetConfigDir(dir string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.configDir = dir
}

func (s *Session) ConfigDir() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.configDir
}

// SetAppContext stores the Wails context (cancellation on app shutdown).
func (s *Session) SetAppContext(ctx context.Context) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.ctx = ctx
}

func (s *Session) Ctx() context.Context {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.ctx
}

func (s *Session) SetClient(c *graphapi.Client, profileName string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.client = c
	s.profileName = profileName
}

// SetTokens attaches the token broker of the connected profile.
func (s *Session) SetTokens(t TokenBroker) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.tokens = t
}

// SetIdentity records who the connection is: the tenant id and whether it
// signs in as the app (client credentials) or as a user (delegated).
func (s *Session) SetIdentity(tenantID string, appOnly bool) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.tenantID, s.appOnly = tenantID, appOnly
}

// TenantID is the connected tenant ("" when disconnected).
func (s *Session) TenantID() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.tenantID
}

// AppOnly reports an app-only (client credentials) connection.
func (s *Session) AppOnly() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.appOnly
}

// Tokens returns the broker for non-Graph resources, or nil when disconnected.
func (s *Session) Tokens() TokenBroker {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.tokens
}

// OnDisconnect registers cleanup for when the tenant connection ends (e.g.
// stopping signed-in PowerShell hosts); fn gets the connection that ended,
// so a late cleanup cannot touch a newer connection.
func (s *Session) OnDisconnect(fn func(prev *graphapi.Client)) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.onDisconnect = append(s.onDisconnect, fn)
}

// Conn is everything one tenant connection consists of.
type Conn struct {
	Client      *graphapi.Client
	ProfileName string
	Tokens      TokenBroker
	TenantID    string
	AppOnly     bool
	// ProfileID and Policy come from a saved profile (empty when ad hoc).
	ProfileID    string
	Policy       Policy
	DelegatedOrg string
}

// Connect installs a connection in one step: nobody sees a part of the new
// connection next to a part of the old one (the policy is in place before
// the client can be used). A connection still in place is ended first,
// through Disconnect, so its hooks run (hosts closed, caches dropped,
// connection-bound operations cancelled) before the new one is visible.
func (s *Session) Connect(c Conn) {
	s.connMu.Lock()
	defer s.connMu.Unlock()
	if s.Connected() {
		s.disconnect()
	}
	c.Policy.AllowedGroups = append([]string(nil), c.Policy.AllowedGroups...)
	s.mu.Lock()
	defer s.mu.Unlock()
	s.client, s.profileName, s.tokens = c.Client, c.ProfileName, c.Tokens
	s.tenantID, s.appOnly = c.TenantID, c.AppOnly
	s.policy, s.profileID, s.delegatedOrg = c.Policy, c.ProfileID, c.DelegatedOrg
}

// Current returns the connection as one consistent snapshot (zero when
// disconnected); read it once instead of field by field when several parts
// must belong together.
func (s *Session) Current() Conn {
	s.mu.RLock()
	defer s.mu.RUnlock()
	p := s.policy
	p.AllowedGroups = append([]string(nil), p.AllowedGroups...)
	return Conn{
		Client: s.client, ProfileName: s.profileName, Tokens: s.tokens,
		TenantID: s.tenantID, AppOnly: s.appOnly,
		ProfileID: s.profileID, Policy: p, DelegatedOrg: s.delegatedOrg,
	}
}

func (s *Session) Disconnect() {
	s.connMu.Lock()
	defer s.connMu.Unlock()
	s.disconnect()
}

func (s *Session) disconnect() {
	s.mu.Lock()
	hooks := append([]func(*graphapi.Client){}, s.onDisconnect...)
	prev := s.client
	defer func() {
		for _, fn := range hooks {
			fn(prev)
		}
	}()
	defer s.mu.Unlock()
	s.client = nil
	s.tokens = nil
	s.tenantID, s.appOnly = "", false
	s.profileName = ""
	s.policy, s.profileID, s.delegatedOrg = Policy{}, "", ""
}

func (s *Session) Client() (*graphapi.Client, error) {
	s.mu.RLock()
	defer s.mu.RUnlock()
	if s.client == nil {
		return nil, ErrNotConnected
	}
	return s.client, nil
}

func (s *Session) Connected() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.client != nil
}

func (s *Session) ProfileName() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.profileName
}

func (s *Session) SetReadOnly(v bool) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.readOnly = v
}

func (s *Session) ReadOnly() bool {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.readOnly
}

// GuardWrite is called before any write that does not name a user or group
// (under a group-scoped profile it is refused; see GuardWriteOn).
func (s *Session) GuardWrite() error {
	if err := s.ceiling(false); err != nil {
		return err
	}
	return s.inScope(nil)
}

// GuardDestructive is a write guard plus typed confirm: the user must type
// the target identifier; the backend re-verifies it (we do not trust the UI alone).
func (s *Session) GuardDestructive(target, confirm string) error {
	if err := s.ceiling(true); err != nil {
		return err
	}
	if err := s.inScope(nil); err != nil {
		return err
	}
	return s.confirm(target, confirm)
}

func (s *Session) confirm(target, confirm string) error {
	if confirm != target {
		return errors.New("confirmation text does not match the target — operation cancelled")
	}
	return nil
}

// Record writes the operation outcome to the audit log. detail must hold no sensitive data.
func (s *Session) Record(action, target, detail string, err error) {
	e := auditlog.Entry{
		Action:  action,
		Target:  target,
		Detail:  detail,
		OK:      err == nil,
		Profile: s.ProfileName(),
	}
	if err != nil {
		e.Error = err.Error()
	}
	s.Audit.Write(e)
}

// SetDelegatedOrg records the customer tenant of a partner (GDAP) profile.
func (s *Session) SetDelegatedOrg(org string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.delegatedOrg = org
}

// DelegatedOrg is the customer tenant a partner profile manages ("" if none).
func (s *Session) DelegatedOrg() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.delegatedOrg
}
