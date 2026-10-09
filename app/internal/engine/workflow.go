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

// stepRef marks which step a planned change belongs to.
const stepRef = "__step"

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
// only (no packs, no reads), inputs the actions know, templates that name
// declared inputs. It returns the workflow's danger: the highest of its steps.
func (e *Engine) ValidateWorkflow(fields []Field, steps []WorkflowStep) (Danger, error) {
	if len(steps) == 0 {
		return "", fmt.Errorf("a workflow needs steps")
	}
	declared := map[string]bool{}
	for _, f := range fields {
		declared[f.Name] = true
	}
	danger := Read
	for i, st := range steps {
		where := fmt.Sprintf("step %d (%s)", i+1, st.Action)
		if strings.HasPrefix(st.Action, "pack.") {
			return "", fmt.Errorf("%s: a workflow can use built-in actions only", where)
		}
		a, err := e.lookup(st.Action)
		if err != nil {
			return "", fmt.Errorf("%s: no such action", where)
		}
		if a.Danger == Read {
			return "", fmt.Errorf("%s: read actions show tables, they cannot be steps", where)
		}
		known := map[string]bool{}
		for _, f := range a.Fields {
			known[f.Name] = true
			if f.Required && f.Default == "" && st.With[f.Name] == "" {
				return "", fmt.Errorf("%s: input %q is required", where, f.Name)
			}
		}
		for k, v := range st.With {
			if !known[k] {
				return "", fmt.Errorf("%s: the action has no input %q", where, k)
			}
			for _, m := range templateRe.FindAllStringSubmatch(v, -1) {
				if !declared[m[1]] {
					return "", fmt.Errorf("%s: {{%s}} is not an input of the workflow", where, m[1])
				}
			}
		}
		if st.When != nil && !declared[st.When.Input] {
			return "", fmt.Errorf("%s: condition on unknown input %q", where, st.When.Input)
		}
		switch st.OnError {
		case "", "stop", "continue":
		default:
			return "", fmt.Errorf("%s: onError is stop or continue", where)
		}
		if dangerRank[a.Danger] > dangerRank[danger] {
			danger = a.Danger
		}
	}
	return danger, nil
}

// WorkflowGate reports the first step that cannot run now.
func (e *Engine) WorkflowGate(steps []WorkflowStep) *Reason {
	for _, st := range steps {
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
		changes, err := w.planStep(env, a, stepIn)
		if err != nil {
			if st.OnError == "continue" {
				out = append(out, Change{Target: st.Action, Field: "step", Op: "none", Note: "stepFailed", After: err.Error(), Step: st.Action,
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
			ch.Step = st.Action
			out = append(out, ch)
		}
	}
	return out, nil
}

func (w workflowImpl) planStep(env Env, a Action, in Inputs) ([]Change, error) {
	if err := validate(a.Manifest, in); err != nil {
		return nil, err
	}
	// The profile's limits apply to every step with its own inputs.
	if err := w.e.checkPolicy(env, a, in); err != nil {
		return nil, err
	}
	impl, reason := w.e.resolve(a)
	if impl == nil {
		return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	return impl.Plan(env, in)
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
	impl, reason := w.e.resolve(a)
	if impl == nil {
		return &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	stepIn := st.inputs(in)
	_ = validate(a.Manifest, stepIn) // fills defaults as at plan time
	return impl.Apply(env, stepIn, ch)
}

// StopOnError tells Apply whether a failed change ends the run.
func (w workflowImpl) StopOnError(ch Change) bool {
	_, st, err := w.step(ch)
	return err != nil || st.OnError != "continue"
}
