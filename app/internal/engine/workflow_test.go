package engine

import (
	"errors"
	"strings"
	"testing"
)

// stepImpl plans one change naming its inputs and can be told to fail.
type stepImpl struct {
	applied *[]string
	fail    bool
}

func (stepImpl) Backend() Backend { return BackendGraph }
func (s stepImpl) Plan(_ Env, in Inputs) ([]Change, error) {
	return []Change{{Target: in["user"], Field: "x", Op: "set", After: in["state"]}}, nil
}
func (s stepImpl) Apply(_ Env, in Inputs, ch Change) error {
	if s.fail {
		return errors.New("boom")
	}
	*s.applied = append(*s.applied, ch.Target+":"+in["state"])
	return nil
}

func TestWorkflowChainsActionsWithTemplatesConditionsAndStop(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{}, WorkflowProvider{})
	user := Field{Name: "user", Kind: FieldUser, Required: true}
	e.Register(
		Action{Manifest: Manifest{ID: "a.block", Danger: Write, Fields: []Field{user, {Name: "state", Kind: FieldText}}}, Impls: []Impl{stepImpl{applied: &applied}}},
		Action{Manifest: Manifest{ID: "a.fails", Danger: Destructive, ConfirmField: "user", Fields: []Field{user}}, Impls: []Impl{stepImpl{applied: &applied, fail: true}}},
		Action{Manifest: Manifest{ID: "a.read", Danger: Read}, Impls: []Impl{ReadImpl(nil)}},
	)
	fields := []Field{user, {Name: "extra", Kind: FieldChoice, Options: []string{"yes", "no"}, Default: "no"}}
	steps := []WorkflowStep{
		{Action: "a.block", With: map[string]string{"user": "{{user}}", "state": "blocked"}},
		{Action: "a.block", With: map[string]string{"user": "{{user}}", "state": "extra"}, When: &StepCondition{Input: "extra", Equals: "yes"}},
		{Action: "a.fails", With: map[string]string{"user": "{{user}}"}},
		{Action: "a.block", With: map[string]string{"user": "{{user}}", "state": "after"}},
	}
	danger, _, err := e.ValidateWorkflow(fields, "user", steps)
	if err != nil || danger != Destructive {
		t.Fatalf("validate: %v %v", danger, err)
	}
	for name, bad := range map[string][]WorkflowStep{
		"unknown action": {{Action: "nope"}},
		"read step":      {{Action: "a.read"}},
		"pack step":      {{Action: "pack.x.y"}},
		"unknown input":  {{Action: "a.block", With: map[string]string{"user": "{{who}}"}}},
		"unknown field":  {{Action: "a.block", With: map[string]string{"user": "x", "color": "red"}}},
		"missing input":  {{Action: "a.block", With: map[string]string{"state": "x"}}},
		"confirm other":  {{Action: "a.fails", With: map[string]string{"user": "{{extra}}"}}},
	} {
		if _, _, err := e.ValidateWorkflow(fields, "user", bad); err == nil {
			t.Errorf("%s must be refused", name)
		}
	}
	e.Register(Action{Manifest: Manifest{ID: "pack.p.wf", Danger: danger, Fields: fields, ConfirmField: "user", Pack: "p", Workflow: true},
		Impls: []Impl{e.NewWorkflow(steps)}})

	p, err := e.Plan("pack.p.wf", Inputs{"user": "ann"})
	if err != nil {
		t.Fatal(err)
	}
	// The conditional step is left out; each change names its step.
	if len(p.Changes) != 3 || p.Changes[0].Step != "a.block" || p.Changes[1].Step != "a.fails" || p.Changes[0].After != "blocked" {
		t.Fatalf("plan %+v", p.Changes)
	}
	res, err := e.Apply(p.ID, "ann")
	if err != nil {
		t.Fatal(err)
	}
	// The failing step stops the run: the last step does not happen.
	if strings.Join(applied, ",") != "ann:blocked" || res.Failed != 2 || !strings.Contains(res.Outcomes[2].Error, "earlier step failed") {
		t.Fatalf("applied %v result %+v", applied, res)
	}

	p, _ = e.Plan("pack.p.wf", Inputs{"user": "bob", "extra": "yes"})
	if len(p.Changes) != 4 || p.Changes[1].After != "extra" {
		t.Fatalf("the condition must include the step: %+v", p.Changes)
	}

	// A step whose backend is missing makes the whole workflow unavailable.
	if r := e.WorkflowGate([]WorkflowStep{{Action: "a.block"}}); r != nil {
		t.Fatalf("gate: %+v", r)
	}
	s.Disconnect()
	if r := e.WorkflowGate(steps); r == nil || r.Key != "stepUnavailable" {
		t.Fatalf("gate when the tenant is gone: %+v", r)
	}
}

func TestWorkflowConfirmationsAndScope(t *testing.T) {
	s := newSession(t)
	var applied []string
	e := New(s, GraphProvider{}, WorkflowProvider{})
	user := Field{Name: "user", Kind: FieldUser, Required: true}
	strong := fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "a", Op: "set", Ref: map[string]string{ConfirmRef: "delete 3"}}}}
	strong2 := fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "b", Op: "set", Ref: map[string]string{ConfirmRef: "delete 5"}}}}
	idle := fakeImpl{backend: BackendGraph, applied: &applied, changes: []Change{{Target: "c", Op: "none", Ref: map[string]string{ConfirmRef: "delete 0"}}}}
	e.Register(
		Action{Manifest: Manifest{ID: "x.one", Danger: Destructive, ConfirmField: "user", Fields: []Field{user}}, Impls: []Impl{strong}},
		Action{Manifest: Manifest{ID: "x.two", Danger: Destructive, ConfirmField: "user", Fields: []Field{user}}, Impls: []Impl{strong2}},
		Action{Manifest: Manifest{ID: "x.idle", Danger: Destructive, ConfirmField: "user", Fields: []Field{user}}, Impls: []Impl{idle}},
	)
	steps := []WorkflowStep{
		{Action: "x.one", With: map[string]string{"user": "{{user}}"}},
		{Action: "x.idle", With: map[string]string{"user": "{{user}}"}},
		{Action: "x.two", With: map[string]string{"user": "{{user}}"}},
	}
	if _, _, err := e.ValidateWorkflow([]Field{user}, "user", steps); err != nil {
		t.Fatal(err)
	}
	e.Register(Action{Manifest: Manifest{ID: "pack.p.x", Danger: Destructive, Fields: []Field{user}, ConfirmField: "user", Pack: "p", Workflow: true},
		Impls: []Impl{e.NewWorkflow(steps)}})
	p, err := e.Plan("pack.p.x", Inputs{"user": "ann"})
	if err != nil {
		t.Fatal(err)
	}
	// Both strong confirmations are asked for; the idle step adds none.
	if p.ConfirmTarget != "delete 3 + delete 5" {
		t.Fatalf("confirm %q", p.ConfirmTarget)
	}
	if _, err := e.Apply(p.ID, "delete 5"); err == nil {
		t.Fatal("only part of the confirmation must not apply")
	}
}
