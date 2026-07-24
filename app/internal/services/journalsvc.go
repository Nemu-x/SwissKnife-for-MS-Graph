package services

import (
	"errors"

	"swissknife-app/internal/journal"
	"swissknife-app/internal/session"
)

// JournalService exposes the persistent run journal to the UI (History page).
type JournalService struct {
	s *session.Session
}

func NewJournalService(s *session.Session) *JournalService { return &JournalService{s: s} }

// List returns the newest runs (bounded), newest first.
func (j *JournalService) List(limit int) ([]journal.RunSummary, error) {
	if j.s.Journal == nil {
		return nil, errors.New("run journal unavailable")
	}
	if limit <= 0 || limit > 100 {
		limit = 100
	}
	return j.s.Journal.List(limit), nil
}

// Get returns one run with all of its journaled events.
func (j *JournalService) Get(opID string) (*journal.Run, error) {
	if j.s.Journal == nil {
		return nil, errors.New("run journal unavailable")
	}
	return j.s.Journal.Get(opID)
}
