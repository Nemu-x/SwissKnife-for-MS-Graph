package services

import (
	"encoding/json"

	"swissknife-app/internal/session"
)

// ServiceHealthService — Microsoft 365 service health and message center.
type ServiceHealthService struct {
	s *session.Session
}

func NewServiceHealthService(s *session.Session) *ServiceHealthService {
	return &ServiceHealthService{s: s}
}

// Overview lists per-service health status.
func (h *ServiceHealthService) Overview() ([]json.RawMessage, error) {
	c, err := h.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(h.s.Ctx(), "/admin/serviceAnnouncement/healthOverviews", nil, 0)
}

// Issues lists active/recent service issues.
func (h *ServiceHealthService) Issues() ([]json.RawMessage, error) {
	c, err := h.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(h.s.Ctx(), "/admin/serviceAnnouncement/issues", nil, 200)
}

// Messages lists Message Center posts (roadmap/changes).
func (h *ServiceHealthService) Messages() ([]json.RawMessage, error) {
	c, err := h.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(h.s.Ctx(), "/admin/serviceAnnouncement/messages", nil, 200)
}
