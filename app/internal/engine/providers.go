package engine

import (
	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/session"
)

// GraphProvider is usable whenever a tenant is connected.
type GraphProvider struct{}

func (GraphProvider) Backend() Backend { return BackendGraph }

func (GraphProvider) Status(s *session.Session) *Reason {
	if !s.Connected() {
		return &Reason{Key: "notConnected"}
	}
	return nil
}

// LDAPProvider reports the on-prem directory connection; with Secure set it
// is the encrypted variant password resets need.
type LDAPProvider struct {
	Secure bool
	Get    func() *ldapx.Client
}

func (p LDAPProvider) Backend() Backend {
	if p.Secure {
		return BackendLDAPTLS
	}
	return BackendLDAP
}

func (p LDAPProvider) Status(*session.Session) *Reason {
	c := p.Get()
	if c == nil {
		return &Reason{Key: "ldapNotConnected"}
	}
	if p.Secure && !c.Secure() {
		return &Reason{Key: "ldapNeedsTLS"}
	}
	return nil
}
