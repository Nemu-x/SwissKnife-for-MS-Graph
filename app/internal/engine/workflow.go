package engine

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"swissknife-app/internal/session"
)

// Workflows chain built-in catalog actions: no code of their own, every
// step planned and applied by its action's implementation, all of it one
// preview, one confirmation and one journal entry.

// BackendWorkflow runs a workflow; its steps pick their own backends.
const BackendWorkflow Backend = "workflow"

// WorkflowProvider is always available: the steps carry the real checks.
type WorkflowProvider struct{}

func (WorkflowProvider) Backend() Backend { return BackendWorkflow }

func (WorkflowProvider) Status(*session.Session) *Reason { return nil }

// WorkflowStep runs one catalog action.
type WorkflowStep struct {
	Action string            // built-in action id
	With   map[string]string // the action's inputs; "{{name}}" takes a workflow input
	// When, if set, runs the step only when the workflow input equals Equals.
	When    *StepCondition
	OnError string // "stop" (default): later steps do not run; "continue"
}

// StepCondition compares a workflow input with a value.
type StepCondition struct {
	Input  string
	Equals string
}

// stepRef marks which step a planned change belongs to; backendRef which
// implementation planned it.
const (
	stepRef    = "__step"
	backendRef = "__backend"
)

var templateRe = regexp.MustCompile(`\{\{\s*([A-Za-z][A-Za-z0-9]*)\s*\}\}`)

func render(tmpl string, in Inputs) string {
	return templateRe.ReplaceAllStringFunc(tmpl, func(m string) string {
		return in[templateRe.FindStringSubmatch(m)[1]]
	})
}

func (st WorkflowStep) runs(in Inputs) bool {
	return st.When == nil || strings.EqualFold(in[st.When.Input], st.When.Equals)
}

func (st WorkflowStep) inputs(in Inputs) Inputs {
	out := Inputs{}
	for k, v := range st.With {
		out[k] = render(v, in)
	}
	return out
}

// ValidateWorkflow checks a workflow against the catalog: built-in actions
// only (no packs, no reads, no on-prem directory, no secret inputs), inputs
// the actions know, templates that name declared inputs, and a confirmation
// that names what every destructive step acts on. It returns the workflow's
// danger (the highest of its steps) and the permissions its steps need.
func (e *Engine) ValidateWorkflow(fields []Field, confirmField string, steps []WorkflowStep) (Danger, []string, error) {
	if len(steps) == 0 {
		return "", nil, fmt.Errorf("a workflow needs steps")
	}
	declared := map[string]bool{}
	confirmKind := FieldKind("")
	for _, f := range fields {
		declared[f.Name] = true
		if f.Name == confirmField {
			confirmKind = f.Kind
		}
	}
	if confirmField != "" && confirmKind != FieldUser && confirmKind != FieldGroup && confirmKind != FieldText {
		return "", nil, fmt.Errorf("confirmField must name a user, group or text field of the workflow")
	}
	danger := Read
	var perms []string
	permSeen := map[string]bool{}
	for i, st := range steps {
		where := fmt.Sprintf("step %d (%s)", i+1, st.Action)
		if strings.HasPrefix(st.Action, "pack.") {
			return "", nil, fmt.Errorf("%s: a workflow can use built-in actions only", where)
		}
		a, err := e.lookup(st.Action)
		if err != nil {
			return "", nil, fmt.Errorf("%s: no such action", where)
		}
		if a.Danger == Read {
			return "", nil, fmt.Errorf("%s: read actions show tables, they cannot be steps", where)
		}
		for _, impl := range a.Impls {
			if usesLDAP(impl.Backend()) {
				return "", nil, fmt.Errorf("%s: on-prem directory actions cannot be workflow steps", where)
			}
		}
		for _, f := range a.Fields {
			if f.Kind == FieldSecret {
				return "", nil, fmt.Errorf("%s: actions that take a password cannot be workflow steps", where)
			}
		}
		// What the operator retypes must be exactly what a destructive step
		// acts on: never one name confirmed while another is changed.
		if a.Danger == Destructive {
			if confirmField == "" {
				return "", nil, fmt.Errorf("%s is destructive: the workflow names its confirmField", where)
			}
			if a.ConfirmField != "" && strings.TrimSpace(st.With[a.ConfirmField]) != "{{"+confirmField+"}}" {
				return "", nil, fmt.Errorf("%s is destructive: its %q must be {{%s}}, the confirmed field", where, a.ConfirmField, confirmField)
			}
		}
		for _, perm := range a.Permissions {
			if !permSeen[perm] {
				permSeen[perm] = true
				perms = append(perms, perm)
			}
		}
		known := map[string]bool{}
		for _, f := range a.Fields {
			known[f.Name] = true
			if f.Required && f.Default == "" && st.With[f.Name] == "" {
				return "", nil, fmt.Errorf("%s: input %q is required", where, f.Name)
			}
		}
		for k, v := range st.With {
			if !known[k] {
				return "", nil, fmt.Errorf("%s: the action has no input %q", where, k)
			}
			for _, m := range templateRe.FindAllStringSubmatch(v, -1) {
				if !declared[m[1]] {
					return "", nil, fmt.Errorf("%s: {{%s}} is not an input of the workflow", where, m[1])
				}
			}
		}
		if st.When != nil && !declared[st.When.Input] {
			return "", nil, fmt.Errorf("%s: condition on unknown input %q", where, st.When.Input)
		}
		switch st.OnError {
		case "", "stop", "continue":
		default:
			return "", nil, fmt.Errorf("%s: onError is stop or continue", where)
		}
		if dangerRank[a.Danger] > dangerRank[danger] {
			danger = a.Danger
		}
	}
	return danger, perms, nil
}

// WorkflowGate reports the first required step that cannot run now; an
// optional step (when, or onError: continue) is reported in the plan instead.
func (e *Engine) WorkflowGate(steps []WorkflowStep) *Reason {
	for _, st := range steps {
		if st.When != nil || st.OnError == "continue" {
			continue
		}
		a, err := e.lookup(st.Action)
		if err != nil {
			return &Reason{Key: "stepUnavailable", Params: map[string]string{"action": st.Action}}
		}
		if impl, r := e.resolve(a); impl == nil {
			p := map[string]string{"action": st.Action}
			if r != nil {
				p["reason"] = r.Key
			}
			return &Reason{Key: "stepUnavailable", Params: p}
		}
	}
	return nil
}

// NewWorkflow is the implementation of a workflow action.
func (e *Engine) NewWorkflow(steps []WorkflowStep) Impl { return workflowImpl{e: e, steps: steps} }

type workflowImpl struct {
	e     *Engine
	steps []WorkflowStep
}

func (w workflowImpl) Backend() Backend { return BackendWorkflow }

func (w workflowImpl) step(ch Change) (int, WorkflowStep, error) {
	i, err := strconv.Atoi(ch.Ref[stepRef])
	if err != nil || i < 0 || i >= len(w.steps) {
		return 0, WorkflowStep{}, fmt.Errorf("unknown workflow step")
	}
	return i, w.steps[i], nil
}

// Plan plans every step that runs, against the live state, and tags each
// change with its step. A step's own stronger confirmation still applies.
func (w workflowImpl) Plan(env Env, in Inputs) ([]Change, error) {
	var out []Change
	for i, st := range w.steps {
		if !st.runs(in) {
			continue
		}
		a, err := w.e.lookup(st.Action)
		if err != nil {
			return nil, err
		}
		stepIn := st.inputs(in)
		changes, backend, err := w.planStep(env, a, stepIn)
		if err != nil {
			if st.OnError == "continue" {
				out = append(out, Change{Target: st.Action, Field: "step", Op: "none", Note: "stepFailed", After: err.Error(), Step: st.Action, StepNo: i + 1,
					Ref: map[string]string{stepRef: strconv.Itoa(i)}})
				continue
			}
			return nil, fmt.Errorf("step %d (%s): %w", i+1, st.Action, err)
		}
		for _, ch := range changes {
			if ch.Ref == nil {
				ch.Ref = map[string]string{}
			}
			ch.Ref[stepRef] = strconv.Itoa(i)
			// Apply uses the implementation that made this change.
			ch.Ref[backendRef] = string(backend)
			ch.Step, ch.StepNo = st.Action, i+1
			out = append(out, ch)
		}
	}
	return out, nil
}

func (w workflowImpl) planStep(env Env, a Action, in Inputs) ([]Change, Backend, error) {
	if err := validate(a.Manifest, in); err != nil {
		return nil, "", err
	}
	// The profile's limits apply to every step with its own inputs.
	if err := w.e.checkPolicy(env, a, in); err != nil {
		return nil, "", err
	}
	impl, reason := w.e.resolve(a)
	if impl == nil {
		return nil, "", &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	changes, err := impl.Plan(env, in)
	return changes, impl.Backend(), err
}

// implFor returns the action's implementation on backend, if it can run now.
func (w workflowImpl) implFor(a Action, backend Backend) (Impl, error) {
	for _, impl := range a.Impls {
		if impl.Backend() != backend {
			continue
		}
		if p, ok := w.e.providers[backend]; ok {
			if r := p.Status(w.e.s); r != nil {
				return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + r.Key}
			}
		}
		return impl, nil
	}
	return nil, errPlanGone
}

func (w workflowImpl) Apply(env Env, in Inputs, ch Change) error {
	_, st, err := w.step(ch)
	if err != nil {
		return err
	}
	a, err := w.e.lookup(st.Action)
	if err != nil {
		return err
	}
	impl, err := w.implFor(a, Backend(ch.Ref[backendRef]))
	if err != nil {
		return err
	}
	stepIn := st.inputs(in)
	_ = validate(a.Manifest, stepIn) // fills defaults as at plan time
	// The profile's limits may have changed since the preview.
	if err := w.e.checkPolicy(env, a, stepIn); err != nil {
		return err
	}
	return impl.Apply(env, stepIn, ch)
}

// StopOnError tells Apply whether a failed change ends the run.
func (w workflowImpl) StopOnError(ch Change) bool {
	_, st, err := w.step(ch)
	return err != nil || st.OnError != "continue"
}
