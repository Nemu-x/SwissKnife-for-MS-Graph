package engine

import (
	"crypto/rand"
	"encoding/hex"
	"sync"
	"time"
)

// planTTL bounds how stale a preview may be when it is applied: the tenant
// can change underneath it, so an old plan must be recomputed.
const planTTL = 15 * time.Minute

// Plan is a stored preview. Apply accepts only its ID, so what runs is
// exactly what the operator saw.
type Plan struct {
	ID            string    `json:"id"`
	ActionID      string    `json:"actionId"`
	Backend       Backend   `json:"backend"`
	Inputs        Inputs    `json:"inputs"`
	Changes       []Change  `json:"changes"`
	ConfirmTarget string    `json:"confirmTarget,omitempty"`
	CreatedAt     time.Time `json:"createdAt"`

	impl Impl
	conn any // connection the plan was computed on (*graphapi.Client or *ldapx.Client)
}

// target is the plan's headline object for journal and audit.
func (p *Plan) target() string {
	if p.ConfirmTarget != "" {
		return p.ConfirmTarget
	}
	if len(p.Changes) > 0 {
		return p.Changes[0].Target
	}
	return ""
}

type planStore struct {
	mu    sync.Mutex
	plans map[string]*Plan
	now   func() time.Time
}

func newPlanStore() *planStore {
	return &planStore{plans: map[string]*Plan{}, now: time.Now}
}

func (s *planStore) put(p *Plan) {
	var b [8]byte
	_, _ = rand.Read(b[:])
	s.mu.Lock()
	defer s.mu.Unlock()
	p.ID = hex.EncodeToString(b[:])
	p.CreatedAt = s.now()
	// Drop expired plans on the way in; previews are small but unbounded.
	for id, old := range s.plans {
		if s.now().Sub(old.CreatedAt) > planTTL {
			delete(s.plans, id)
		}
	}
	s.plans[p.ID] = p
}

// get returns a live plan without consuming it.
func (s *planStore) get(id string) (*Plan, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	p, ok := s.plans[id]
	if !ok || s.now().Sub(p.CreatedAt) > planTTL {
		return nil, errPlanGone
	}
	return p, nil
}

func (s *planStore) drop(id string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	delete(s.plans, id)
}

// take removes and returns a live plan: a plan is applied at most once.
func (s *planStore) take(id string) (*Plan, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	p, ok := s.plans[id]
	if !ok {
		return nil, errPlanGone
	}
	delete(s.plans, id)
	if s.now().Sub(p.CreatedAt) > planTTL {
		return nil, errPlanGone
	}
	return p, nil
}
