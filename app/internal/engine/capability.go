package engine

import (
	"sort"

	"swissknife-app/internal/session"
)

// A capability is the stable public name of what an action does
// ("exchange.mailbox.fullAccess"), independent of the action's id and of the
// implementations behind it. Workflows may name steps by capability; the
// Capabilities view shows, per capability, which implementation runs now,
// which would take over, and why the others cannot.

// Router is implemented by providers whose backend can run on this machine
// or on the paired worker; Via reports which ("local" | "worker").
type Router interface {
	Via(s *session.Session) string
}

// ImplStatus is one implementation of a capability, in preference order.
type ImplStatus struct {
	Backend Backend `json:"backend"`
	// State: "runs" (the one the resolver picks), "ready" (usable fallback)
	// or "unavailable" (Reason says why).
	State  string  `json:"state"`
	Via    string  `json:"via,omitempty"`
	Reason *Reason `json:"reason,omitempty"`
}

// CapabilityView is one capability with its implementation chain.
type CapabilityView struct {
	Capability string            `json:"capability"`
	Action     string            `json:"action"`
	Page       string            `json:"page"`
	Danger     Danger            `json:"danger"`
	Pack       string            `json:"pack,omitempty"`
	Label      map[string]string `json:"label,omitempty"`
	Impls      []ImplStatus      `json:"impls"`
	// Reason is set when the action as a whole cannot run (an untrusted
	// pack, the profile's limits), whatever its implementations.
	Reason *Reason `json:"reason,omitempty"`
}

func capabilityOf(m Manifest) string {
	if m.Capability != "" {
		return m.Capability
	}
	return m.ID
}

// CapabilityAction returns the id of the action that provides capability c.
func (e *Engine) CapabilityAction(c string) (string, bool) {
	e.mu.RLock()
	defer e.mu.RUnlock()
	for _, a := range e.actions {
		if capabilityOf(a.Manifest) == c {
			return a.ID, true
		}
	}
	return "", false
}

// Capabilities lists every capability with its implementation chain.
func (e *Engine) Capabilities() []CapabilityView {
	e.mu.RLock()
	all := make([]Action, 0, len(e.actions))
	for _, a := range e.actions {
		all = append(all, a)
	}
	e.mu.RUnlock()
	out := make([]CapabilityView, 0, len(all))
	for _, a := range all {
		v := CapabilityView{Capability: capabilityOf(a.Manifest), Action: a.ID, Page: a.Page, Danger: a.Danger,
			Pack: a.Pack, Label: a.Label, Impls: []ImplStatus{}}
		if a.Gate != nil {
			v.Reason = a.Gate()
		}
		if v.Reason == nil && !onDirectory(a) && !dangerAllowed(effectiveDanger(a), e.s.Policy().MaxDanger) {
			v.Reason = &Reason{Key: "policy"}
		}
		picked := false
		for _, impl := range a.Impls {
			st := ImplStatus{Backend: impl.Backend(), State: "unavailable"}
			p, ok := e.providers[impl.Backend()]
			if !ok {
				st.Reason = &Reason{Key: "backendMissing", Params: map[string]string{"backend": string(impl.Backend())}}
			} else if st.Reason = p.Status(e.s); st.Reason == nil {
				st.State = "ready"
				if !picked && v.Reason == nil {
					st.State, picked = "runs", true
				}
				if r, ok := p.(Router); ok {
					st.Via = r.Via(e.s)
				}
			}
			v.Impls = append(v.Impls, st)
		}
		out = append(out, v)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Capability < out[j].Capability })
	return out
}
