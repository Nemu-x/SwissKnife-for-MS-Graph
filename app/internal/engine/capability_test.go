package engine

import (
	"testing"

	"swissknife-app/internal/session"
)

type routedProvider struct{ fakeProvider }

func (routedProvider) Via(*session.Session) string { return "worker" }

func TestCapabilitiesShowTheChain(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s,
		fakeProvider{backend: "exo-api", reason: &Reason{Key: "notEnabled"}},
		routedProvider{fakeProvider{backend: "pwsh"}},
		fakeProvider{backend: "graph"},
	)
	e.Register(Action{
		Manifest: Manifest{ID: "a", Capability: "exchange.thing.set", Danger: Write},
		Impls: []Impl{
			fakeImpl{backend: "exo-api", applied: &applied},
			fakeImpl{backend: "pwsh", applied: &applied},
			fakeImpl{backend: "graph", applied: &applied},
			fakeImpl{backend: "nowhere", applied: &applied},
		},
	}, Action{Manifest: Manifest{ID: "b", Danger: Write}, Impls: []Impl{fakeImpl{backend: "graph", applied: &applied}}})

	if id, ok := e.CapabilityAction("exchange.thing.set"); !ok || id != "a" {
		t.Fatalf("CapabilityAction = %q %v", id, ok)
	}
	if id, ok := e.CapabilityAction("b"); !ok || id != "b" {
		t.Fatalf("an action without a name is its own capability: %q %v", id, ok)
	}
	caps := e.Capabilities()
	if len(caps) != 2 || caps[1].Capability != "exchange.thing.set" {
		t.Fatalf("caps = %+v", caps)
	}
	got := caps[1].Impls
	want := []struct{ state, via, reason string }{
		{"unavailable", "", "notEnabled"}, {"runs", "worker", ""}, {"ready", "", ""}, {"unavailable", "", "backendMissing"},
	}
	for i, w := range want {
		r := ""
		if got[i].Reason != nil {
			r = got[i].Reason.Key
		}
		if got[i].State != w.state || got[i].Via != w.via || r != w.reason {
			t.Errorf("impl %d = %+v, want %+v", i, got[i], w)
		}
	}
}

func TestDuplicateCapabilityPanics(t *testing.T) {
	e := New(newSession(t))
	e.Register(Action{Manifest: Manifest{ID: "a", Capability: "x.y.z"}, Impls: []Impl{fakeImpl{backend: "graph"}}})
	defer func() {
		if recover() == nil {
			t.Fatal("a second action with the same capability registered")
		}
	}()
	e.Register(Action{Manifest: Manifest{ID: "b", Capability: "x.y.z"}, Impls: []Impl{fakeImpl{backend: "graph"}}})
}

// Under a group scope, tenant-wide changes and pack scripts can never pass
// the policy: the catalog and the capabilities view say so up front.
func TestGroupScopeMarksUnscopedActions(t *testing.T) {
	s := newSession(t)
	s.SetPolicy("", session.Policy{AllowedGroups: []string{"g1"}})
	e := New(s, fakeProvider{backend: "graph"})
	e.Register(
		Action{Manifest: Manifest{ID: "tenantWide", Danger: Write}, Impls: []Impl{fakeImpl{backend: "graph"}}},
		Action{Manifest: Manifest{ID: "perUser", Danger: Write, Fields: []Field{{Name: "user", Kind: FieldUser}}}, Impls: []Impl{fakeImpl{backend: "graph"}}},
		Action{Manifest: Manifest{ID: "report", Danger: Read}, Impls: []Impl{fakeImpl{backend: "graph"}}},
	)
	want := map[string]string{"tenantWide": "policyUnscoped", "perUser": "", "report": ""}
	for _, c := range e.Catalog() {
		got := ""
		if c.Reason != nil {
			got = c.Reason.Key
		}
		if got != want[c.ID] {
			t.Errorf("catalog %s: reason %q, want %q", c.ID, got, want[c.ID])
		}
	}
	for _, c := range e.Capabilities() {
		runs := c.Impls[0].State == "runs"
		if runs != (want[c.Action] == "") {
			t.Errorf("capabilities %s: %+v", c.Action, c)
		}
	}
}
