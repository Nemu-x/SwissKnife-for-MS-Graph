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
	"sync"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// Backend names a transport an implementation runs on.
type Backend string

const (
	BackendGraph Backend = "graph"
	// On-premises Active Directory; LDAPTLS is the same connection when it
	// is encrypted (password resets need it).
	BackendLDAP    Backend = "ldap"
	BackendLDAPTLS Backend = "ldap-tls"
)

// usesLDAP reports a backend that runs against the on-prem directory, not
// the tenant.
func usesLDAP(b Backend) bool { return b == BackendLDAP || b == BackendLDAPTLS }

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
	// FieldSnapshot picks one of the saved configuration snapshots.
	FieldSnapshot FieldKind = "snapshot"
	// FieldSecret is a password input: never echoed back in a plan.
	FieldSecret FieldKind = "secret"
)

// Field is one input of an action. Labels are i18n keys on the frontend
// (actions.fields.<name>, actions.options.<option>), never text here.
type Field struct {
	Name     string    `json:"name"`
	Kind     FieldKind `json:"kind"`
	Required bool      `json:"required"`
	Options  []string  `json:"options,omitempty"` // FieldChoice
	Default  string    `json:"default,omitempty"`
	// Label by language, for actions from packs (built-in ones use i18n).
	Label map[string]string `json:"label,omitempty"`
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
	// Label, Hint and Pack describe actions from community packs; built-in
	// actions are labelled by the UI from i18n.
	Label map[string]string `json:"label,omitempty"`
	Hint  map[string]string `json:"hint,omitempty"`
	Pack  string            `json:"pack,omitempty"`
	// Workflow marks an action that chains other actions (from a pack).
	Workflow bool `json:"workflow,omitempty"`
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
	// Step is the action id of the workflow step this change belongs to.
	Step string `json:"step,omitempty"`
}

// ConfirmRef is the Change.Ref key an implementation sets to replace the
// typed confirmation of a destructive action.
const ConfirmRef = "confirm"

// Env is what an implementation may use while planning or applying.
type Env struct {
	Ctx    context.Context
	Graph  *graphapi.Client
	Tokens session.TokenBroker // non-Graph resources; nil before connect
	// TenantID and AppOnly describe the connection (Exchange needs both:
	// the tenant in its URLs, the sign-in kind for its routing hint).
	TenantID string
	// DelegatedOrg is the customer tenant a partner (GDAP) profile manages.
	DelegatedOrg string
	// LDAP is the active on-prem directory connection (nil if none).
	LDAP *ldapx.Client
	// ExchangeOrg / ExchangeUPN, when set, are what Exchange PowerShell
	// signs in with, so no Graph call is needed to find them (a remote
	// worker gets them from the client instead of a Graph token).
	ExchangeOrg string
	ExchangeUPN string
	AppOnly     bool
	// PS runs PowerShell cmdlets for the PowerShell backends; nil when the
	// engine has none.
	PS PSRunner
	// ReadOnly is the session's read-only switch, for the rare preview that
	// must write something itself (an eDiscovery search) to compute a plan.
	ReadOnly bool
}

// PSRunner runs one cmdlet in the PowerShell host of a module family
// ("exo", "teams", "ipps"), connecting it for env's connection first.
type PSRunner interface {
	Invoke(env Env, family, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error)
}

// ScriptRunner runs a trusted action-pack script (the pwsh pool implements it).
type ScriptRunner interface {
	RunScript(env Env, family, script string, params map[string]any) ([]json.RawMessage, error)
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
	// Gate, when set, can make the action unavailable (an untrusted pack).
	Gate func() *Reason
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
	mu        sync.RWMutex // guards actions (packs are swapped at run time)
	s         *session.Session
	actions   map[string]Action
	providers map[Backend]Provider
	plans     *planStore
	// PS is handed to implementations through Env.
	PS PSRunner
	// Grants returns the Graph permissions the connection holds, or nil when
	// unknown; the catalog warns about the ones a manifest lacks.
	Grants func() map[string]bool
	// LDAP returns the active on-prem directory connection (nil if none).
	LDAP func() *ldapx.Client
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
	e.mu.Lock()
	defer e.mu.Unlock()
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
	// MissingPermissions are Graph permissions the token does not carry: the
	// action stays available (a policy may still allow it) but is likely to
	// fail with 403.
	MissingPermissions []string `json:"missingPermissions,omitempty"`
	// FanOut: a read that can run across several tenants (Graph only).
	FanOut bool `json:"fanOut,omitempty"`
}

// Catalog lists every action, sorted by id for a stable UI.
func (e *Engine) Catalog() []CatalogEntry {
	var have map[string]bool
	if e.Grants != nil && e.s.Connected() {
		have = e.Grants()
	}
	e.mu.RLock()
	all := make([]Action, 0, len(e.actions))
	for _, a := range e.actions {
		all = append(all, a)
	}
	e.mu.RUnlock()
	out := make([]CatalogEntry, 0, len(all))
	for _, a := range all {
		entry := CatalogEntry{Manifest: a.Manifest, FanOut: graphReader(a) != nil}
		if entry.Fields == nil {
			entry.Fields = []Field{} // the UI maps over it: [] not null
		}
		impl, reason := e.resolve(a)
		switch {
		case !onDirectory(a) && !dangerAllowed(effectiveDanger(a), e.s.Policy().MaxDanger):
			entry.Reason = &Reason{Key: "policy"}
		case impl != nil:
			entry.Available, entry.Backend = true, impl.Backend()
			entry.MissingPermissions = missingPermissions(a.Manifest, have)
		default:
			entry.Reason = reason
		}
		out = append(out, entry)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].ID < out[j].ID })
	return out
}

// SetPacks replaces every action from community packs (ids "pack.…").
func (e *Engine) SetPacks(actions []Action) {
	e.mu.Lock()
	defer e.mu.Unlock()
	for id := range e.actions {
		if strings.HasPrefix(id, "pack.") {
			delete(e.actions, id)
		}
	}
	for _, a := range actions {
		if !strings.HasPrefix(a.ID, "pack.") || len(a.Impls) == 0 {
			continue
		}
		if _, dup := e.actions[a.ID]; !dup {
			e.actions[a.ID] = a
		}
	}
}

// resolve returns the first implementation whose backend is usable, or the
// reason the most preferred one is not.
func (e *Engine) resolve(a Action) (Impl, *Reason) {
	if a.Gate != nil {
		if r := a.Gate(); r != nil {
			return nil, r
		}
	}
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

// Env is the environment implementations get, for callers outside the
// catalog that reuse the engine's connections (snapshot collectors).
func (e *Engine) Env(ctx context.Context) Env { return e.env(ctx) }

// BackendStatus reports whether a backend can run now (nil = available).
func (e *Engine) BackendStatus(b Backend) *Reason {
	p, ok := e.providers[b]
	if !ok {
		return &Reason{Key: "backendMissing", Params: map[string]string{"backend": string(b)}}
	}
	return p.Status(e.s)
}

func (e *Engine) env(ctx context.Context) Env {
	c, _ := e.s.Client()
	var dir *ldapx.Client
	if e.LDAP != nil {
		dir = e.LDAP()
	}
	return Env{Ctx: ctx, Graph: c, Tokens: e.s.Tokens(), TenantID: e.s.TenantID(), DelegatedOrg: e.s.DelegatedOrg(), LDAP: dir, AppOnly: e.s.AppOnly(), PS: e.PS, ReadOnly: e.s.ReadOnly()}
}

func (e *Engine) lookup(id string) (Action, error) {
	e.mu.RLock()
	a, ok := e.actions[id]
	e.mu.RUnlock()
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
	if err := e.checkPolicy(env, a, in); err != nil {
		return nil, err
	}
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
		conn:     connOf(impl.Backend(), env),
	}
	if a.ConfirmField != "" {
		p.ConfirmTarget = in[a.ConfirmField]
	}
	// An implementation may ask for a stronger confirmation than the input
	// (e.g. the sender plus the number of messages a purge will delete).
	for _, ch := range changes {
		if c := ch.Ref[ConfirmRef]; c != "" {
			p.ConfirmTarget = c
		}
	}
	e.plans.put(p)
	// The caller gets a copy: Apply must run the stored preview, not one a
	// caller edited after the fact.
	out := *p
	out.Inputs = cloneInputs(p.Inputs)
	// Secrets (a new password) stay in the stored plan only.
	for _, f := range a.Fields {
		if f.Kind == FieldSecret && out.Inputs[f.Name] != "" {
			out.Inputs[f.Name] = "********"
		}
	}
	out.Changes = make([]Change, len(p.Changes))
	for i, ch := range p.Changes {
		if ch.Ref != nil {
			ref := make(map[string]string, len(ch.Ref))
			for k, v := range ch.Ref {
				ref[k] = v
			}
			ch.Ref = ref
		}
		out.Changes[i] = ch
	}
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
	// A pack may have lost its trust since the preview.
	if a.Gate != nil {
		if r := a.Gate(); r != nil {
			e.plans.drop(planID)
			return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + r.Key}
		}
	}
	// The plan holds object ids of the connection it was computed on (the
	// tenant, or the on-prem directory): a disconnect or a switch invalidates it.
	cur := e.env(e.s.Ctx())
	if usesLDAP(p.Backend) {
		if cur.LDAP == nil {
			return nil, &Error{Code: "unavailable", Msg: "action unavailable: ldapNotConnected"}
		}
	} else if cur.Graph == nil {
		return nil, session.ErrNotConnected
	}
	if connOf(p.Backend, cur) != p.conn {
		e.plans.drop(planID)
		return nil, errPlanGone
	}
	if prov, ok := e.providers[p.Backend]; ok {
		if reason := prov.Status(e.s); reason != nil {
			return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
		}
	}
	// The profile's limits may have changed since the preview.
	if err := e.checkPolicy(e.env(e.s.Ctx()), a, p.Inputs); err != nil {
		return nil, err
	}
	switch {
	case usesLDAP(p.Backend):
		// A tenant profile's limits say nothing about the on-prem directory:
		// only the read-only switch and the typed confirmation apply.
		if err := e.s.GuardDirectory(a.Danger == Destructive, p.ConfirmTarget, confirm); err != nil {
			return nil, err
		}
	case a.Danger == Destructive:
		if err := e.s.GuardDestructiveChecked(p.ConfirmTarget, confirm); err != nil {
			return nil, err
		}
	case a.Danger == Write:
		if err := e.s.GuardWriteChecked(); err != nil {
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
	res := &Result{OpID: op.ID, Outcomes: []Outcome{}}
	var firstErr error
	// A workflow step that fails can end the run: later changes do not run.
	stopper, _ := p.impl.(interface{ StopOnError(Change) bool })
	stopped := false
	for _, ch := range p.Changes {
		out := Outcome{Change: ch, OK: true}
		switch {
		case ch.Op == "none":
			res.Skipped++
			out.Skipped = true
		case stopped:
			out.OK, out.Error = false, "not run: an earlier step failed"
			res.Failed++
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
				if stopper != nil && stopper.StopOnError(ch) {
					stopped = true
				}
			} else {
				res.Applied++
			}
		}
		res.Outcomes = append(res.Outcomes, out)
		if j := e.s.Journal; j != nil {
			j.Event(op.ID, "item", map[string]any{
				"step":   ch.Step,
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
	detail := fmt.Sprintf("backend=%s applied=%d skipped=%d failed=%d %s", p.Backend, res.Applied, res.Skipped, res.Failed, describe(p.Changes))
	if usesLDAP(p.Backend) && env.LDAP != nil {
		// The session's profile is a tenant's; name the directory too.
		detail = "directory=" + env.LDAP.Config().Name + " " + detail
	}
	e.s.Record("action."+a.ID, target, detail, firstErr)
	return res, nil
}

// Execute plans and applies an action in one go on ctx, for multi-step runs
// (playbooks) that already show their own preview, hold their own operation
// and journal, and have passed the destructive confirmation for the whole
// run. Only the write guard is checked here. It returns the first
// implementation error as is (not display-wrapped) so the caller can read
// its structure.
func (e *Engine) Execute(ctx context.Context, actionID string, in Inputs) (*Result, error) {
	a, err := e.lookup(actionID)
	if err != nil {
		return nil, err
	}
	in = cloneInputs(in)
	if err := validate(a.Manifest, in); err != nil {
		return nil, err
	}
	if a.Danger != Read {
		if err := e.s.GuardWriteChecked(); err != nil {
			return nil, err
		}
	}
	impl, reason := e.resolve(a)
	if impl == nil {
		return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	env := e.env(ctx)
	if err := e.checkPolicy(env, a, in); err != nil {
		return nil, err
	}
	changes, err := impl.Plan(env, in)
	if err != nil {
		return nil, err
	}
	res := &Result{Outcomes: []Outcome{}}
	var firstErr error
	for _, ch := range changes {
		out := Outcome{Change: ch, OK: true}
		switch {
		case ch.Op == "none":
			res.Skipped++
			out.Skipped = true
		case ctx.Err() != nil:
			res.Canceled, res.Failed = true, res.Failed+1
			out.OK, out.Error = false, "canceled"
		default:
			if err := impl.Apply(env, in, ch); err != nil {
				out.OK, out.Error = false, e.WrapErr(err).Error()
				res.Failed++
				if firstErr == nil {
					firstErr = err
				}
			} else {
				res.Applied++
			}
		}
		res.Outcomes = append(res.Outcomes, out)
	}
	e.s.Record("action."+a.ID, (&Plan{Changes: changes}).target(),
		fmt.Sprintf("backend=%s applied=%d skipped=%d failed=%d %s", impl.Backend(), res.Applied, res.Skipped, res.Failed, describe(changes)),
		firstErr)
	return res, firstErr
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

// connOf is the connection a plan for backend b is bound to.
func connOf(b Backend, env Env) any {
	if usesLDAP(b) {
		return env.LDAP
	}
	return env.Graph
}

// onDirectory reports an action that runs only against the on-prem directory.
func onDirectory(a Action) bool {
	for _, impl := range a.Impls {
		if !usesLDAP(impl.Backend()) {
			return false
		}
	}
	return len(a.Impls) > 0
}

// ActionManifest returns a registered action's manifest.
func (e *Engine) ActionManifest(id string) (Manifest, bool) {
	a, err := e.lookup(id)
	return a.Manifest, err == nil
}
