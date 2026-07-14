package services

import (
	"encoding/json"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/session"
)

type AuditService struct {
	s *session.Session
}

func NewAuditService(s *session.Session) *AuditService { return &AuditService{s: s} }

func (a *AuditService) DirectoryAudits(top int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	return c.ListAll(a.s.Ctx(), "/auditLogs/directoryAudits", topParams(top), top)
}

func (a *AuditService) SignIns(top int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	return c.ListAll(a.s.Ctx(), "/auditLogs/signIns", topParams(top), top)
}

// Activity is the local app action log (ADR-002).
func (a *AuditService) Activity(n int) []auditlog.Entry {
	if n <= 0 {
		n = 200
	}
	return a.s.Audit.Tail(n)
}
