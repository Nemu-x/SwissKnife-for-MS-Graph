package services

import (
	"context"
	"errors"
	"fmt"
	"strings"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/secrets"
)

// fanOutLimit bounds the tenants of one run; each is read in turn.
const fanOutLimit = 25

// RunAcross runs a read action in every chosen saved profile, one after the
// other, and merges the rows under a leading "tenant" column. Only app-only
// profiles take part (a delegated one would need a sign-in per tenant);
// tenants that fail are named in the note, the others still answer. The run
// is an operation: CancelAcross or a disconnect stops it.
func (a *ActionsService) RunAcross(actionID string, inputs map[string]string, profileIDs []string) (res *engine.ReadResult, err error) {
	var names, failed []string
	defer func() {
		detail := fmt.Sprintf("profiles=%s failed=%d", strings.Join(names, ","), len(failed))
		if res != nil {
			detail += fmt.Sprintf(" rows=%d", len(res.Rows))
		}
		a.s.Record("action.fanout", actionID, detail, err)
	}()
	if a.store == nil {
		return nil, errors.New("profiles are not available")
	}
	if len(profileIDs) == 0 {
		return nil, errors.New("choose at least one profile")
	}
	if len(profileIDs) > fanOutLimit {
		return nil, fmt.Errorf("at most %d profiles at a time", fanOutLimit)
	}
	if err := a.e.CheckFanOut(actionID, inputs); err != nil {
		return nil, engineErr(err)
	}
	op, err := a.s.Ops.Start(a.s.Ctx(), ops.KindFanOut)
	if err != nil {
		return nil, err
	}
	defer a.s.Ops.Finish(op)
	ctx, cancel := context.WithTimeout(op.Ctx, 30*time.Minute)
	defer cancel()

	out := &engine.ReadResult{Backend: engine.BackendGraph}
	for _, id := range profileIDs {
		name, r, err := a.runIn(ctx, id, actionID, inputs)
		names = append(names, name)
		if ctx.Err() != nil {
			return nil, errors.New("stopped")
		}
		if err != nil {
			failed = append(failed, name+" ("+shortErr(err)+")")
			continue
		}
		if out.Columns == nil {
			out.Columns = append([]string{"tenant"}, r.Columns...)
		}
		for _, row := range r.Rows {
			tagged := engine.Row{"tenant": name}
			for k, v := range row {
				tagged[k] = v
			}
			out.Rows = append(out.Rows, tagged)
		}
		// A tenant's note (an incomplete scan, a section it may not read)
		// must not get lost in the merge.
		if r.Note != nil {
			out.TenantNotes = append(out.TenantNotes, engine.TenantNote{Tenant: name, Note: *r.Note})
		}
	}
	if len(failed) == len(profileIDs) {
		return nil, errors.New("no tenant answered: " + strings.Join(failed, "; "))
	}
	if out.Columns == nil {
		out.Columns = []string{"tenant"}
	}
	if len(failed) > 0 {
		out.Note = &engine.Reason{Key: "tenantsFailed", Params: map[string]string{"list": strings.Join(failed, "; ")}}
	}
	return out, nil
}

// CancelAcross stops a running cross-tenant read.
func (a *ActionsService) CancelAcross() { a.s.Ops.CancelKind(ops.KindFanOut) }

// shortErr is one readable line of an error: the message of a structured
// Graph error, the first line of a sign-in failure.
func shortErr(err error) string {
	msg := err.Error()
	var oe *OpError
	if errors.As(err, &oe) {
		msg = oe.Message
		if oe.Hint != "" {
			msg += " — " + oe.Hint
		}
	}
	if i := strings.IndexAny(msg, "\r\n"); i > 0 {
		msg = msg[:i]
	}
	if len(msg) > 160 {
		msg = msg[:160] + "…"
	}
	return msg
}

// runIn connects to one saved profile just for this read.
func (a *ActionsService) runIn(ctx context.Context, profileID, actionID string, inputs map[string]string) (string, *engine.ReadResult, error) {
	ctx, cancel := context.WithTimeout(ctx, 3*time.Minute)
	defer cancel()
	name, env, err := profileEnv(ctx, a.store, profileID)
	if name == "" {
		name = profileID
	}
	if err != nil {
		return name, nil, err
	}
	res, err := a.e.RunOn(env, actionID, inputs)
	return name, res, err
}

// profileEnv builds a read-only Graph environment for a saved app-only
// profile (a variable so tests can stand in for the sign-in).
var profileEnv = func(ctx context.Context, store *secrets.Store, profileID string) (string, engine.Env, error) {
	cr, err := ResolveCredentials(store, ConnectRequest{ProfileID: profileID})
	if err != nil {
		return profileID, engine.Env{}, err
	}
	if cr.AuthMode == string(auth.ModeDeviceCode) {
		return cr.Name, engine.Env{}, errors.New("delegated profile: connect to it directly")
	}
	provider, err := NewTokenProvider(cr, nil)
	if err != nil {
		return cr.Name, engine.Env{}, err
	}
	return cr.Name, engine.Env{Ctx: ctx, Graph: graphapi.New(provider), Tokens: provider, TenantID: cr.TenantID, AppOnly: true, ReadOnly: true}, nil
}

// FanOutProfiles lists the saved profiles a cross-tenant read can use.
func (a *ActionsService) FanOutProfiles() ([]secrets.Profile, error) {
	if a.store == nil {
		return []secrets.Profile{}, nil
	}
	list, err := a.store.List()
	out := []secrets.Profile{}
	for _, p := range list {
		if p.AuthMode != string(auth.ModeDeviceCode) {
			out = append(out, p)
		}
	}
	return out, err
}
