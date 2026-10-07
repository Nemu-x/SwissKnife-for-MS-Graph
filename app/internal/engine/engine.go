// Package engine is the execution layer (ADR-008): a catalog of actions, each
// with one or more backend implementations, a resolver that picks the first
// implementation whose backend is usable right now, and a Plan → Apply
// contract so every write is previewed before it touches the tenant.
//
// The engine owns the cross-cutting rules — read-only and typed-confirm
// guards, single-flight ops, the run journal and the audit log — so an action
// implementation only describes what to read (Plan) and how to write one
// change (Apply).
package engine

import (
	"context"
	"encoding/json"
	"fmt"
	"sort"
	"strings"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// Backend names a transport an implementation runs on.
type Backend string

const (
	BackendGraph Backend = "graph"
)

// Danger decides which guards an action passes through.
type Danger string

const (
	Read        Danger = "read"
	Write       Danger = "write"
	Destructive Danger = "destructive" // typed confirm of Manifest.ConfirmField
)

// FieldKind tells the UI which input to render.
type FieldKind string

const (
	FieldUser   FieldKind = "user"
	FieldGroup  FieldKind = "group"
	FieldSku    FieldKind = "sku"
	FieldText   FieldKind = "text"
	FieldChoice FieldKind = "choice"
)

// Field is one input of an action. Labels are i18n keys on the frontend
// (actions.fields.<name>, actions.options.<option>), never text here.
type Field struct {
	Name     string    `json:"name"`
	Kind     FieldKind `json:"kind"`
	Required bool      `json:"required"`
	Options  []string  `json:"options,omitempty"` // FieldChoice
	Default  string    `json:"default,omitempty"`
}

// Manifest describes an action independently of how it runs.
type Manifest struct {
	ID     string  `json:"id"`
	Page   string  `json:"page"` // page whose tiles show the action
	Danger Danger  `json:"danger"`
	Fields []Field `json:"fields"`
	// ConfirmField names the input the operator must retype to apply a
	// destructive action (the backend re-checks it).
	ConfirmField string `json:"confirmField,omitempty"`
	// Permissions the implementations need, for the preflight hint.
	Permissions []string `json:"permissions,omitempty"`
}

// Inputs are the operator's field values, keyed by Field.Name.
type Inputs map[string]string

// Change is one planned modification. Op is add | remove | set | none; "none"
// means the target is already in the requested state and Apply skips it.
// Before/After are display values (names or i18n value tokens).
type Change struct {
	Target string `json:"target"`
	Field  string `json:"field"`
	Op     string `json:"op"`
	Before string `json:"before,omitempty"`
	After  string `json:"after,omitempty"`
	// Note is an i18n token explaining a "none" or unusual change
	// (actions.notes.<note>), e.g. why an inherited license stays.
	Note string `json:"note,omitempty"`
	// Ref carries implementation data Apply needs (resolved ids); not shown.
	Ref map[string]string `json:"-"`
}

// Env is what an implementation may use while planning or applying.
type Env struct {
	Ctx    context.Context
	Graph  *graphapi.Client
	Tokens session.TokenBroker // non-Graph resources; nil before connect
	// TenantID and AppOnly describe the connection (Exchange needs both:
	// the tenant in its URLs, the sign-in kind for its routing hint).
	TenantID string
	AppOnly  bool
	// PS runs PowerShell cmdlets for the PowerShell backends; nil when the
	// engine has none.
	PS PSRunner
}

// PSRunner runs one cmdlet in the PowerShell host of a module family
// ("exo", "teams", "ipps"), connecting it for env's connection first.
type PSRunner interface {
	Invoke(env Env, family, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error)
}

// Impl is one way to run an action on one backend.
type Impl interface {
	Backend() Backend
	// Plan reads the live state and returns the changes; it must not write.
	Plan(env Env, in Inputs) ([]Change, error)
	// Apply performs a single non-"none" change from the plan.
	Apply(env Env, in Inputs, ch Change) error
}

// Action is a manifest plus its implementations in preference order.
type Action struct {
	Manifest
	Impls []Impl
}

// Reason explains why an action is unavailable: an i18n key plus params.
type Reason struct {
	Key    string            `json:"key"`
	Params map[string]string `json:"params,omitempty"`
}

// Provider reports whether a backend is usable for the current session.
type Provider interface {
	Backend() Backend
	Status(s *session.Session) *Reason // nil = available
}

// Error carries a stable code next to the English message so the UI can
// translate engine refusals.
type Error struct {
	Code string
	Msg  string
}

func (e *Error) Error() string { return e.Msg }

// Engine wires actions, providers and plans to one session.
type Engine struct {
	s         *session.Session
	actions   map[string]Action
	providers map[Backend]Provider
	plans     *planStore
	// PS is handed to implementations through Env.
	PS PSRunner
	// WrapErr converts implementation errors for display (services sets the
	// operr envelope with permission hints); identity by default.
	WrapErr func(error) error
}

func New(s *session.Session, providers ...Provider) *Engine {
	e := &Engine{
		s:         s,
		actions:   map[string]Action{},
		providers: map[Backend]Provider{},
		plans:     newPlanStore(),
		WrapErr:   func(err error) error { return err },
	}
	for _, p := range providers {
		e.providers[p.Backend()] = p
	}
	return e
}

// Register adds actions; a duplicate id is a programming error.
func (e *Engine) Register(actions ...Action) {
	for _, a := range actions {
		if _, dup := e.actions[a.ID]; dup {
			panic("engine: duplicate action id " + a.ID)
		}
		if len(a.Impls) == 0 {
			panic("engine: action without implementations " + a.ID)
		}
		e.actions[a.ID] = a
	}
}

// CatalogEntry is a manifest with its availability right now.
type CatalogEntry struct {
	Manifest
	Available bool    `json:"available"`
	Backend   Backend `json:"backend,omitempty"` // implementation that would run
	Reason    *Reason `json:"reason,omitempty"`
}

// Catalog lists every action, sorted by id for a stable UI.
func (e *Engine) Catalog() []CatalogEntry {
	out := make([]CatalogEntry, 0, len(e.actions))
	for _, a := range e.actions {
		entry := CatalogEntry{Manifest: a.Manifest}
		impl, reason := e.resolve(a)
		if impl != nil {
			entry.Available, entry.Backend = true, impl.Backend()
		} else {
			entry.Reason = reason
		}
		out = append(out, entry)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].ID < out[j].ID })
	return out
}

// resolve returns the first implementation whose backend is usable, or the
// reason the most preferred one is not.
func (e *Engine) resolve(a Action) (Impl, *Reason) {
	var first *Reason
	for _, impl := range a.Impls {
		reason := &Reason{Key: "backendMissing", Params: map[string]string{"backend": string(impl.Backend())}}
		if p, ok := e.providers[impl.Backend()]; ok {
			reason = p.Status(e.s)
		}
		if reason == nil {
			return impl, nil
		}
		if first == nil {
			first = reason
		}
	}
	return nil, first
}

func (e *Engine) env(ctx context.Context) Env {
	c, _ := e.s.Client()
	return Env{Ctx: ctx, Graph: c, Tokens: e.s.Tokens(), TenantID: e.s.TenantID(), AppOnly: e.s.AppOnly(), PS: e.PS}
}

func (e *Engine) lookup(id string) (Action, error) {
	a, ok := e.actions[id]
	if !ok {
		return Action{}, &Error{Code: "unknownAction", Msg: "unknown action " + id}
	}
	return a, nil
}

func validate(m Manifest, in Inputs) error {
	for _, f := range m.Fields {
		v := in[f.Name]
		if v == "" && f.Default != "" {
			in[f.Name] = f.Default
			v = f.Default
		}
		if f.Required && v == "" {
			return &Error{Code: "missingInput", Msg: "missing input: " + f.Name}
		}
		if f.Kind == FieldChoice && v != "" && !contains(f.Options, v) {
			return &Error{Code: "badInput", Msg: fmt.Sprintf("invalid value %q for %s", v, f.Name)}
		}
	}
	return nil
}

func contains(list []string, v string) bool {
	for _, x := range list {
		if x == v {
			return true
		}
	}
	return false
}

// Plan previews an action. Previewing is allowed in read-only mode — only
// Apply writes.
func (e *Engine) Plan(actionID string, in Inputs) (*Plan, error) {
	a, err := e.lookup(actionID)
	if err != nil {
		return nil, err
	}
	in = cloneInputs(in)
	if err := validate(a.Manifest, in); err != nil {
		return nil, err
	}
	impl, reason := e.resolve(a)
	if impl == nil {
		return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	env := e.env(e.s.Ctx())
	changes, err := impl.Plan(env, in)
	if err != nil {
		return nil, e.WrapErr(err)
	}
	p := &Plan{
		ActionID: a.ID,
		Backend:  impl.Backend(),
		Inputs:   in,
		Changes:  changes,
		impl:     impl,
		client:   env.Graph,
	}
	if a.ConfirmField != "" {
		p.ConfirmTarget = in[a.ConfirmField]
	}
	e.plans.put(p)
	// The caller gets a copy: Apply must run the stored preview, not one a
	// caller edited after the fact.
	out := *p
	out.Inputs = cloneInputs(p.Inputs)
	out.Changes = append([]Change(nil), p.Changes...)
	return &out, nil
}

// Apply executes a stored plan. confirm must equal the plan's ConfirmTarget
// for destructive actions.
//
// A refused apply (read-only, wrong confirm, another apply running) leaves the
// plan in place so the operator can retry; the plan is consumed only once the
// run actually starts.
func (e *Engine) Apply(planID, confirm string) (*Result, error) {
	p, err := e.plans.get(planID)
	if err != nil {
		return nil, err
	}
	a, err := e.lookup(p.ActionID)
	if err != nil {
		return nil, err
	}
	// The plan holds object ids of the tenant it was computed on: a
	// disconnect or a switch to another profile invalidates it.
	c, err := e.s.Client()
	if err != nil {
		return nil, err
	}
	if c != p.client {
		e.plans.drop(planID)
		return nil, errPlanGone
	}
	if prov, ok := e.providers[p.Backend]; ok {
		if reason := prov.Status(e.s); reason != nil {
			return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
		}
	}
	switch a.Danger {
	case Destructive:
		if err := e.s.GuardDestructive(p.ConfirmTarget, confirm); err != nil {
			return nil, err
		}
	case Write:
		if err := e.s.GuardWrite(); err != nil {
			return nil, err
		}
	}

	op, err := e.s.Ops.Start(e.s.Ctx(), ops.KindAction)
	if err != nil {
		return nil, err
	}
	defer e.s.Ops.Finish(op)
	if _, err := e.plans.take(planID); err != nil {
		return nil, err // applied concurrently or expired meanwhile
	}

	target := p.target()
	if j := e.s.Journal; j != nil {
		j.Begin(op.ID, map[string]any{
			"kind": "action", "action": a.ID, "target": target,
			"backend": string(p.Backend), "changes": p.Changes,
		})
	}

	env := e.env(op.Ctx)
	res := &Result{OpID: op.ID}
	var firstErr error
	for _, ch := range p.Changes {
		out := Outcome{Change: ch, OK: true}
		switch {
		case ch.Op == "none":
			res.Skipped++
			out.Skipped = true
		case op.Canceled():
			res.Canceled = true
			out.OK, out.Error = false, "canceled"
			res.Failed++
		default:
			if err := p.impl.Apply(env, p.Inputs, ch); err != nil {
				werr := e.WrapErr(err)
				out.OK, out.Error = false, werr.Error()
				res.Failed++
				if firstErr == nil {
					firstErr = err
				}
			} else {
				res.Applied++
			}
		}
		res.Outcomes = append(res.Outcomes, out)
		if j := e.s.Journal; j != nil {
			j.Event(op.ID, "item", map[string]any{
				"target": ch.Target, "field": ch.Field, "op": ch.Op,
				"before": ch.Before, "after": ch.After,
				"ok": out.OK, "skipped": out.Skipped, "error": out.Error,
			})
		}
	}

	if j := e.s.Journal; j != nil {
		j.End(op.ID, map[string]any{
			"ok": res.Failed == 0, "applied": res.Applied, "skipped": res.Skipped,
			"failed": res.Failed, "canceled": res.Canceled,
		})
	}
	e.s.Record("action."+a.ID, target,
		fmt.Sprintf("backend=%s applied=%d skipped=%d failed=%d %s", p.Backend, res.Applied, res.Skipped, res.Failed, describe(p.Changes)),
		firstErr)
	return res, nil
}

// Cancel aborts a running apply by op id.
func (e *Engine) Cancel(opID string) { e.s.Ops.Cancel(opID) }

// Outcome is the result of one planned change.
type Outcome struct {
	Change
	OK      bool   `json:"ok"`
	Skipped bool   `json:"skipped"`
	Error   string `json:"error,omitempty"`
}

// Result summarises an apply.
type Result struct {
	OpID     string    `json:"opId"`
	Outcomes []Outcome `json:"outcomes"`
	Applied  int       `json:"applied"`
	Skipped  int       `json:"skipped"`
	Failed   int       `json:"failed"`
	Canceled bool      `json:"canceled"`
}

// describe renders the planned changes for the audit detail, e.g.
// "ann@contoso.com signIn set allowed→blocked". Values are display values
// (names, tokens) — never secrets.
func describe(changes []Change) string {
	parts := make([]string, 0, len(changes))
	for _, c := range changes {
		parts = append(parts, fmt.Sprintf("%s %s %s %s→%s", c.Target, c.Field, c.Op, c.Before, c.After))
	}
	return strings.Join(parts, "; ")
}

func cloneInputs(in Inputs) Inputs {
	out := make(Inputs, len(in))
	for k, v := range in {
		out[k] = v
	}
	return out
}

// errPlanGone is returned for unknown, consumed or expired plans.
var errPlanGone = &Error{Code: "planExpired", Msg: "the preview has expired or was already applied — preview again"}
