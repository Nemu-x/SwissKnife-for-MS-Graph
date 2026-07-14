// Package session holds the current connection and safety guards (ADR-002):
// read-only mode, typed confirm for destructive operations, audit-log recording.
package session

import (
	"context"
	"errors"
	"sync"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
)

var ErrNotConnected = errors.New("not connected — connect to a tenant first")
var ErrReadOnly = errors.New("read-only mode is enabled — writes are blocked")

type Session struct {
	mu          sync.RWMutex
	ctx         context.Context
	client      *graphapi.Client
	profileName string
	readOnly    bool

	Audit *auditlog.Log
}

func New(audit *auditlog.Log) *Session {
	return &Session{ctx: context.Background(), Audit: audit}
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

func (s *Session) Disconnect() {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.client = nil
	s.profileName = ""
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

// GuardWrite is called before any write operation.
func (s *Session) GuardWrite() error {
	if s.ReadOnly() {
		return ErrReadOnly
	}
	return nil
}

// GuardDestructive is a write guard plus typed confirm: the user must type
// the target identifier; the backend re-verifies it (we do not trust the UI alone).
func (s *Session) GuardDestructive(target, confirm string) error {
	if err := s.GuardWrite(); err != nil {
		return err
	}
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
