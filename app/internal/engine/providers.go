package engine

import "swissknife-app/internal/session"

// GraphProvider is usable whenever a tenant is connected.
type GraphProvider struct{}

func (GraphProvider) Backend() Backend { return BackendGraph }

func (GraphProvider) Status(s *session.Session) *Reason {
	if !s.Connected() {
		return &Reason{Key: "notConnected"}
	}
	return nil
}
