package engine

import (
	"encoding/json"
	"errors"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

func TestProfilePolicyCeilingAndGroupScope(t *testing.T) {
	// ann is in the helpdesk scope group g1; bob is not.
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "GET" && r.URL.Path == "/users" {
			upn := strings.TrimSuffix(strings.TrimPrefix(r.URL.Query().Get("$filter"), "userPrincipalName eq '"), "'")
			_ = json.NewEncoder(w).Encode(map[string]any{"value": []map[string]string{{"id": "id-" + strings.Split(upn, "@")[0]}}})
			return
		}
		var body struct {
			GroupIDs []string `json:"groupIds"`
		}
		_ = json.NewDecoder(r.Body).Decode(&body)
		if r.URL.Path == "/users/id-ann/checkMemberGroups" || r.URL.Path == "/groups/g-nested/checkMemberGroups" {
			_ = json.NewEncoder(w).Encode(map[string]any{"value": body.GroupIDs[:1]})
			return
		}
		_, _ = w.Write([]byte(`{"value":[]}`))
	}))
	t.Cleanup(srv.Close)
	s := newSession(t)
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")

	var applied []string
	e := New(s, GraphProvider{})
	impl := fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "x", Op: "set"}}}
	e.Register(
		Action{Manifest: Manifest{ID: "w", Danger: Write, Fields: []Field{{Name: "user", Kind: FieldUser}, {Name: "group", Kind: FieldGroup}}}, Impls: []Impl{impl}},
		Action{Manifest: Manifest{ID: "d", Danger: Destructive, ConfirmField: "user", Fields: []Field{{Name: "user", Kind: FieldUser, Required: true}}}, Impls: []Impl{impl}},
	)

	// Ceiling: write allows w, hides and refuses d — and the session guard
	// refuses destructive runs outside the catalog too.
	s.SetPolicy("", session.Policy{MaxDanger: "write"})
	for _, c := range e.Catalog() {
		if c.ID == "d" && (c.Available || c.Reason == nil || c.Reason.Key != "policy") {
			t.Fatalf("d must be unavailable by policy: %+v", c)
		}
	}
	var ee *Error
	if _, err := e.Plan("d", Inputs{"user": "ann@contoso.com"}); !errors.As(err, &ee) || ee.Code != "policyDanger" {
		t.Fatalf("plan d: %v", err)
	}
	if err := s.GuardDestructive("x", "x"); !errors.Is(err, session.ErrPolicyDestructive) {
		t.Fatalf("guard: %v", err)
	}
	s.SetPolicy("", session.Policy{MaxDanger: "read"})
	if err := s.GuardWrite(); !errors.Is(err, session.ErrPolicyRead) {
		t.Fatalf("read-only profile must refuse writes: %v", err)
	}

	// Scope: members of allowed groups only, groups listed or nested.
	s.SetPolicy("", session.Policy{AllowedGroups: []string{"g1"}})
	if _, err := e.Plan("w", Inputs{"user": "ann@contoso.com"}); err != nil {
		t.Fatalf("ann is in scope: %v", err)
	}
	if _, err := e.Plan("w", Inputs{"user": "bob@contoso.com"}); !errors.As(err, &ee) || ee.Code != "policyScope" || !strings.Contains(ee.Msg, "bob") {
		t.Fatalf("bob is out of scope: %v", err)
	}
	if _, err := e.Plan("w", Inputs{"group": "g1"}); err != nil {
		t.Fatalf("an allowed group itself is in scope: %v", err)
	}
	if _, err := e.Plan("w", Inputs{"group": "g-nested"}); err != nil {
		t.Fatalf("a nested group is in scope: %v", err)
	}
	if _, err := e.Plan("w", Inputs{"group": "g-other"}); err == nil {
		t.Fatal("another group is out of scope")
	}
	if _, err := e.Plan("w", Inputs{}); !errors.As(err, &ee) || ee.Code != "policyUnscoped" {
		t.Fatalf("a write naming nobody is tenant-wide: %v", err)
	}
	// A plan made before the scope was set is re-checked on apply.
	s.SetPolicy("", session.Policy{})
	p, err := e.Plan("w", Inputs{"user": "bob@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	s.SetPolicy("", session.Policy{AllowedGroups: []string{"g1"}})
	if _, err := e.Apply(p.ID, ""); err == nil {
		t.Fatal("apply must re-check the scope")
	}
	// Service writes: unscoped ones are refused, named in-scope ones pass.
	s.SetScopeCheck(func(tg session.Target) error { return e.CheckTarget(e.Env(s.Ctx()), FieldKind(tg.Kind), tg.ID) })
	if err := s.GuardWrite(); !errors.Is(err, session.ErrPolicyUnscoped) {
		t.Fatalf("unscoped write: %v", err)
	}
	if err := s.GuardWriteOn(session.User("ann@contoso.com")); err != nil {
		t.Fatalf("ann: %v", err)
	}
	if err := s.GuardWriteOn(session.User("ann@contoso.com"), session.User("bob@contoso.com")); err == nil {
		t.Fatal("every target must be in scope")
	}
	if _, err := e.Execute(s.Ctx(), "w", Inputs{"user": "bob@contoso.com"}); err == nil {
		t.Fatal("playbook steps are scoped too")
	}
	s.Disconnect()
	if p := s.Policy(); p.MaxDanger != "" || len(p.AllowedGroups) != 0 {
		t.Fatalf("disconnect must drop the policy: %+v", p)
	}
}
