package session

import (
	"encoding/json"
)

// Policy is a guard rail a connection profile carries for shared or junior
// use: the highest danger level allowed and, optionally, the groups whose
// members (and which groups) writes may target. It narrows what the app
// does; it is not a security boundary against the app registration.
type Policy struct {
	MaxDanger     string   `json:"maxDanger,omitempty"`     // "" (no limit) | read | write | destructive
	AllowedGroups []string `json:"allowedGroups,omitempty"` // group object ids
	// GroupLabels are display names for the UI only; never used to decide.
	GroupLabels map[string]string `json:"groupLabels,omitempty"`
}

// Scoped reports whether writes are limited to allowed groups.
func (p Policy) Scoped() bool { return len(p.AllowedGroups) > 0 }

// PolicyError is a refusal by the profile policy. It crosses the Wails
// boundary in the same "operr:" envelope as services.OpError, so the UI can
// translate it by code (errors.engine.<code>).
type PolicyError struct {
	Code string
	Msg  string
}

func (e *PolicyError) Error() string {
	b, _ := json.Marshal(map[string]string{"code": e.Code, "message": e.Msg})
	return "operr:" + string(b)
}

var (
	ErrPolicyRead        = &PolicyError{"policyRead", "this connection profile only allows reading"}
	ErrPolicyDestructive = &PolicyError{"policyDanger", "this connection profile does not allow destructive actions"}
	// ErrPolicyUnscoped refuses a write that cannot name its target while the
	// profile is limited to groups (fail closed).
	ErrPolicyUnscoped = &PolicyError{"policyUnscoped", "this connection profile may only change members of its groups, and this operation cannot be checked against them"}
)

// Target is a user or group a write changes, for the group scope.
type Target struct {
	Kind string // "user" | "group"
	ID   string // object id, UPN or group id
}

// User and Group build targets.
func User(v string) Target  { return Target{"user", v} }
func Group(v string) Target { return Target{"group", v} }

// ScopeFunc checks one target against the allowed groups.
type ScopeFunc func(t Target) error

// SetPolicy installs the connected profile's id and policy.
func (s *Session) SetPolicy(profileID string, p Policy) {
	s.mu.Lock()
	defer s.mu.Unlock()
	p.AllowedGroups = append([]string(nil), p.AllowedGroups...)
	s.policy, s.profileID = p, profileID
}

// Policy returns a copy of the connected profile's policy.
func (s *Session) Policy() Policy {
	s.mu.RLock()
	defer s.mu.RUnlock()
	p := s.policy
	p.AllowedGroups = append([]string(nil), p.AllowedGroups...)
	return p
}

// ProfileID is the connected saved profile ("" for ad-hoc connections).
func (s *Session) ProfileID() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.profileID
}

// SetScopeCheck installs the group-scope checker (the engine's).
func (s *Session) SetScopeCheck(fn ScopeFunc) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.scope = fn
}

func (s *Session) ceiling(destructive bool) error {
	if s.ReadOnly() {
		return ErrReadOnly
	}
	switch s.Policy().MaxDanger {
	case "read":
		return ErrPolicyRead
	case "write":
		if destructive {
			return ErrPolicyDestructive
		}
	}
	return nil
}

func (s *Session) inScope(targets []Target) error {
	if !s.Policy().Scoped() {
		return nil
	}
	if len(targets) == 0 {
		return ErrPolicyUnscoped
	}
	s.mu.RLock()
	fn := s.scope
	s.mu.RUnlock()
	if fn == nil {
		return ErrPolicyUnscoped
	}
	for _, t := range targets {
		if t.ID == "" {
			continue
		}
		if err := fn(t); err != nil {
			return err
		}
	}
	return nil
}

// GuardWriteOn is GuardWrite for a write that names what it changes: under a
// group scope every target must be in scope.
func (s *Session) GuardWriteOn(targets ...Target) error {
	if err := s.ceiling(false); err != nil {
		return err
	}
	return s.inScope(targets)
}

// GuardDestructiveOn is GuardDestructive with the group scope of targets.
func (s *Session) GuardDestructiveOn(target, confirm string, targets ...Target) error {
	if err := s.ceiling(true); err != nil {
		return err
	}
	if err := s.inScope(targets); err != nil {
		return err
	}
	return s.confirm(target, confirm)
}

// GuardDangerous is the ceiling of a destructive operation that checks its
// own confirmation word (bulk deletes); it cannot be scoped to groups.
func (s *Session) GuardDangerous() error {
	if err := s.ceiling(true); err != nil {
		return err
	}
	return s.inScope(nil)
}

// GuardWriteChecked and GuardDestructiveChecked are for callers that already
// applied the group scope themselves (the engine, playbooks): ceiling only.
func (s *Session) GuardWriteChecked() error { return s.ceiling(false) }

func (s *Session) GuardDestructiveChecked(target, confirm string) error {
	if err := s.ceiling(true); err != nil {
		return err
	}
	return s.confirm(target, confirm)
}
