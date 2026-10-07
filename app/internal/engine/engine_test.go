package engine

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/session"
)

// fakeImpl records applied changes; plan returns the configured changes.
type fakeImpl struct {
	backend Backend
	changes []Change
	failOn  string
	applied *[]string
}

func (f fakeImpl) Backend() Backend { return f.backend }
func (f fakeImpl) Plan(Env, Inputs) ([]Change, error) {
	return append([]Change(nil), f.changes...), nil
}
func (f fakeImpl) Apply(_ Env, _ Inputs, ch Change) error {
	if ch.Target == f.failOn {
		return errors.New("boom")
	}
	*f.applied = append(*f.applied, ch.Target)
	return nil
}

type fakeProvider struct {
	backend Backend
	reason  *Reason
}

func (p fakeProvider) Backend() Backend                { return p.backend }
func (p fakeProvider) Status(*session.Session) *Reason { return p.reason }

func newSession(t *testing.T) *session.Session {
	t.Helper()
	dir := t.TempDir()
	s := session.New(auditlog.New(dir))
	s.SetJournal(journal.New(filepath.Join(dir, "runs")))
	s.SetClient(graphapi.New(graphapi.StaticToken("t")), "test")
	return s
}

func TestResolverPicksFirstAvailableImpl(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s,
		fakeProvider{backend: "exo-api", reason: &Reason{Key: "notEnabled"}},
		fakeProvider{backend: "pwsh"},
	)
	e.Register(Action{
		Manifest: Manifest{ID: "a", Danger: Write},
		Impls: []Impl{
			fakeImpl{backend: "exo-api", applied: &applied},
			fakeImpl{backend: "pwsh", applied: &applied},
		},
	})
	cat := e.Catalog()
	if len(cat) != 1 || !cat[0].Available || cat[0].Backend != "pwsh" {
		t.Fatalf("catalog = %+v, want available via pwsh", cat)
	}
}

func TestUnavailableReportsMostPreferredReason(t *testing.T) {
	s := newSession(t)
	e := New(s, fakeProvider{backend: "pwsh", reason: &Reason{Key: "moduleMissing"}})
	e.Register(Action{
		Manifest: Manifest{ID: "a", Danger: Read},
		Impls:    []Impl{fakeImpl{backend: "pwsh"}, fakeImpl{backend: "worker"}},
	})
	cat := e.Catalog()
	if cat[0].Available || cat[0].Reason == nil || cat[0].Reason.Key != "moduleMissing" {
		t.Fatalf("catalog = %+v, want unavailable with moduleMissing", cat[0])
	}
	if _, err := e.Plan("a", Inputs{}); err == nil {
		t.Fatal("plan of an unavailable action must fail")
	}
}

func TestValidateRequiredAndChoices(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	e.Register(Action{
		Manifest: Manifest{ID: "a", Danger: Write, Fields: []Field{
			{Name: "user", Kind: FieldUser, Required: true},
			{Name: "op", Kind: FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
		}},
		Impls: []Impl{fakeImpl{backend: BackendGraph, applied: &applied}},
	})
	if _, err := e.Plan("a", Inputs{}); err == nil || !strings.Contains(err.Error(), "user") {
		t.Fatalf("missing user: err = %v", err)
	}
	if _, err := e.Plan("a", Inputs{"user": "u", "op": "drop"}); err == nil {
		t.Fatal("invalid choice must be rejected")
	}
	p, err := e.Plan("a", Inputs{"user": "u"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Inputs["op"] != "add" {
		t.Fatalf("default not applied: %+v", p.Inputs)
	}
}

func TestApplyIsOneShotSkipsNoneAndJournals(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	e.Register(Action{
		Manifest: Manifest{ID: "a", Danger: Write},
		Impls: []Impl{fakeImpl{backend: BackendGraph, applied: &applied, failOn: "c", changes: []Change{
			{Target: "a", Op: "add"}, {Target: "b", Op: "none"}, {Target: "c", Op: "remove"},
		}}},
	})
	p, err := e.Plan("a", Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	r, err := e.Apply(p.ID, "")
	if err != nil {
		t.Fatal(err)
	}
	if r.Applied != 1 || r.Skipped != 1 || r.Failed != 1 || len(applied) != 1 || applied[0] != "a" {
		t.Fatalf("result = %+v applied = %v", r, applied)
	}
	if _, err := e.Apply(p.ID, ""); err == nil {
		t.Fatal("a plan must apply only once")
	}
	runs := s.Journal.List(10)
	if len(runs) != 1 || runs[0].Kind != "action" || runs[0].Summary["failed"] != float64(1) {
		t.Fatalf("journal = %+v", runs)
	}
	audit, _ := os.ReadFile(s.Audit.Path())
	if !strings.Contains(string(audit), "action.a") {
		t.Fatalf("audit missing action entry: %s", audit)
	}
}

func TestApplyGuards(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	impl := fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "x", Op: "set"}}}
	e.Register(
		Action{Manifest: Manifest{ID: "w", Danger: Write}, Impls: []Impl{impl}},
		Action{Manifest: Manifest{ID: "d", Danger: Destructive, ConfirmField: "user",
			Fields: []Field{{Name: "user", Kind: FieldUser, Required: true}}}, Impls: []Impl{impl}},
	)

	s.SetReadOnly(true)
	p, err := e.Plan("w", Inputs{})
	if err != nil {
		t.Fatalf("preview must work in read-only mode: %v", err)
	}
	if _, err := e.Apply(p.ID, ""); !errors.Is(err, session.ErrReadOnly) {
		t.Fatalf("read-only apply: err = %v", err)
	}
	s.SetReadOnly(false)

	p, _ = e.Plan("d", Inputs{"user": "ann@contoso.com"})
	if _, err := e.Apply(p.ID, "bob@contoso.com"); err == nil {
		t.Fatal("wrong confirm must be refused")
	}
	p, _ = e.Plan("d", Inputs{"user": "ann@contoso.com"})
	if _, err := e.Apply(p.ID, "ann@contoso.com"); err != nil {
		t.Fatalf("right confirm: %v", err)
	}
	if len(applied) != 1 {
		t.Fatalf("applied = %v", applied)
	}
}

func TestExpiredPlanIsRefused(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	e.Register(Action{Manifest: Manifest{ID: "a", Danger: Write}, Impls: []Impl{fakeImpl{backend: BackendGraph, applied: &applied}}})
	p, _ := e.Plan("a", Inputs{})
	e.plans.now = func() time.Time { return time.Now().Add(planTTL + time.Minute) }
	_, err := e.Apply(p.ID, "")
	var ee *Error
	if !errors.As(err, &ee) || ee.Code != "planExpired" {
		t.Fatalf("err = %v, want planExpired", err)
	}
}

func TestDuplicateRegistrationPanics(t *testing.T) {
	s := newSession(t)
	e := New(s)
	a := Action{Manifest: Manifest{ID: "a"}, Impls: []Impl{fakeImpl{}}}
	e.Register(a)
	defer func() {
		if recover() == nil {
			t.Fatal("duplicate id must panic")
		}
	}()
	e.Register(a)
}

func TestNotConnectedMakesGraphUnavailable(t *testing.T) {
	s := newSession(t)
	s.Disconnect()
	e := New(s, GraphProvider{})
	e.Register(Action{Manifest: Manifest{ID: "a"}, Impls: []Impl{fakeImpl{backend: BackendGraph}}})
	if c := e.Catalog()[0]; c.Available || c.Reason.Key != "notConnected" {
		t.Fatalf("catalog = %+v", c)
	}
}

func TestRefusedApplyKeepsThePlan(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	e.Register(Action{Manifest: Manifest{ID: "a", Danger: Write},
		Impls: []Impl{fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "x", Op: "set"}}}}})
	p, _ := e.Plan("a", Inputs{})
	s.SetReadOnly(true)
	if _, err := e.Apply(p.ID, ""); !errors.Is(err, session.ErrReadOnly) {
		t.Fatalf("err = %v", err)
	}
	s.SetReadOnly(false)
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatalf("retry after a refusal must work: %v", err)
	}
}

func TestConnectionChangeInvalidatesPlan(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{})
	e.Register(Action{Manifest: Manifest{ID: "a", Danger: Write},
		Impls: []Impl{fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "x", Op: "set"}}}}})

	p, _ := e.Plan("a", Inputs{})
	s.Disconnect()
	if _, err := e.Apply(p.ID, ""); !errors.Is(err, session.ErrNotConnected) {
		t.Fatalf("disconnected apply: err = %v", err)
	}

	// Reconnected to another tenant: the plan's object ids mean nothing there.
	s.SetClient(graphapi.New(graphapi.StaticToken("other")), "other")
	var ee *Error
	if _, err := e.Apply(p.ID, ""); !errors.As(err, &ee) || ee.Code != "planExpired" {
		t.Fatalf("apply after reconnect: err = %v", err)
	}
	if len(applied) != 0 {
		t.Fatalf("nothing may run: %v", applied)
	}
}
