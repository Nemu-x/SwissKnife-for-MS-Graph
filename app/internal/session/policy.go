package session

import "errors"

// Policy is a guard rail a connection profile carries for shared or junior
// use: the highest danger level allowed and, optionally, the groups whose
// members (and which groups) write actions may target. It narrows what the
// app does; it is not a security boundary against the app registration.
type Policy struct {
	MaxDanger     string   `json:"maxDanger,omitempty"`     // "" (no limit) | read | write | destructive
	AllowedGroups []string `json:"allowedGroups,omitempty"` // group object ids
	// GroupLabels are display names for the UI only; never used to decide.
	GroupLabels map[string]string `json:"groupLabels,omitempty"`
}

var (
	ErrPolicyRead        = errors.New("this connection profile only allows reading")
	ErrPolicyDestructive = errors.New("this connection profile does not allow destructive actions")
)

// SetPolicy installs the connected profile's policy.
func (s *Session) SetPolicy(p Policy) {
	s.mu.Lock()
	defer s.mu.Unlock()
	p.AllowedGroups = append([]string(nil), p.AllowedGroups...)
	s.policy = p
}

// Policy returns a copy of the connected profile's policy.
func (s *Session) Policy() Policy {
	s.mu.RLock()
	defer s.mu.RUnlock()
	p := s.policy
	p.AllowedGroups = append([]string(nil), p.AllowedGroups...)
	return p
}
