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
	"swissknife-app/internal/secrets"
)

// fanOutLimit bounds the tenants of one run; each is read in turn.
const fanOutLimit = 25

// RunAcross runs a read action in every chosen saved profile, one after the
// other, and merges the rows under a leading "tenant" column. Only app-only
// profiles take part (a delegated one would need a sign-in per tenant);
// tenants that fail are named in the note, the others still answer.
func (a *ActionsService) RunAcross(actionID string, inputs map[string]string, profileIDs []string) (*engine.ReadResult, error) {
	if a.store == nil {
		return nil, errors.New("profiles are not available")
	}
	if len(profileIDs) == 0 {
		return nil, errors.New("choose at least one profile")
	}
	if len(profileIDs) > fanOutLimit {
		return nil, fmt.Errorf("at most %d profiles at a time", fanOutLimit)
	}
	out := &engine.ReadResult{Backend: engine.BackendGraph}
	var failed []string
	ctx := a.s.Ctx()
	for _, id := range profileIDs {
		name, res, err := a.runIn(ctx, id, actionID, inputs)
		if err != nil {
			var ee *engine.Error
			if errors.As(err, &ee) && ee.Code == "notFanOut" {
				return nil, engineErr(err) // the same for every tenant
			}
			if ctx.Err() != nil {
				return nil, ctx.Err()
			}
			failed = append(failed, fmt.Sprintf("%s (%v)", name, err))
			continue
		}
		if out.Columns == nil {
			out.Columns = append([]string{"tenant"}, res.Columns...)
		}
		for _, r := range res.Rows {
			row := engine.Row{"tenant": name}
			for k, v := range r {
				row[k] = v
			}
			out.Rows = append(out.Rows, row)
		}
	}
	if out.Columns == nil {
		out.Columns = []string{"tenant"}
	}
	if len(failed) > 0 {
		out.Note = &engine.Reason{Key: "tenantsFailed", Params: map[string]string{"list": strings.Join(failed, "; ")}}
	}
	a.s.Record("action.fanout", actionID, fmt.Sprintf("profiles=%d failed=%d rows=%d", len(profileIDs), len(failed), len(out.Rows)), nil)
	return out, nil
}

// runIn connects to one saved profile just for this read.
func (a *ActionsService) runIn(ctx context.Context, profileID, actionID string, inputs map[string]string) (string, *engine.ReadResult, error) {
	ctx, cancel := context.WithTimeout(ctx, 3*time.Minute)
	defer cancel()
	name, env, err := profileEnv(ctx, a.store, profileID)
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
