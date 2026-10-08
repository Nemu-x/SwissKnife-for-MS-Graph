package engine

import (
	"fmt"
	"net/url"
	"strings"
)

var dangerRank = map[Danger]int{Read: 0, Write: 1, Destructive: 2}

// dangerAllowed reports whether a profile ceiling ("" = none) admits d.
func dangerAllowed(d Danger, max string) bool {
	if max == "" {
		return true
	}
	limit, ok := dangerRank[Danger(max)]
	if !ok {
		return false // an unknown ceiling allows nothing risky
	}
	return dangerRank[d] <= limit
}

// checkPolicy applies the connected profile's policy to an action about to
// be planned: the danger ceiling, then (for changes) the group scope of
// every user and group the inputs name.
func (e *Engine) checkPolicy(env Env, a Action, in Inputs) error {
	pol := e.s.Policy()
	if !dangerAllowed(a.Danger, pol.MaxDanger) {
		return &Error{Code: "policyDanger", Msg: fmt.Sprintf("this connection profile allows %s actions at most", pol.MaxDanger)}
	}
	if a.Danger == Read || len(pol.AllowedGroups) == 0 {
		return nil
	}
	for _, f := range a.Fields {
		v := strings.TrimSpace(in[f.Name])
		if v == "" || (f.Kind != FieldUser && f.Kind != FieldGroup) {
			continue
		}
		if err := e.inScope(env, f.Kind, v, pol.AllowedGroups); err != nil {
			return err
		}
	}
	return nil
}

// CheckTarget applies the group scope to one user or group outside the
// catalog (playbooks naming their target).
func (e *Engine) CheckTarget(env Env, kind FieldKind, v string) error {
	pol := e.s.Policy()
	if len(pol.AllowedGroups) == 0 || v == "" {
		return nil
	}
	return e.inScope(env, kind, v, pol.AllowedGroups)
}

// inScope: a user must be a (transitive) member of an allowed group; a group
// must be one of them or nested in one. Lookup failures refuse (fail closed).
func (e *Engine) inScope(env Env, kind FieldKind, v string, allowed []string) error {
	coll := "/users/"
	if kind == FieldGroup {
		coll = "/groups/"
		for _, g := range allowed {
			if strings.EqualFold(g, v) {
				return nil
			}
		}
	}
	if env.Graph == nil {
		return &Error{Code: "policyScope", Msg: "not connected"}
	}
	for start := 0; start < len(allowed); start += 20 { // checkMemberGroups takes 20 ids
		end := min(start+20, len(allowed))
		var out struct {
			Value []string `json:"value"`
		}
		body := map[string]any{"groupIds": allowed[start:end]}
		if err := env.Graph.Post(env.Ctx, coll+url.PathEscape(v)+"/checkMemberGroups", body, &out); err != nil {
			return &Error{Code: "policyScope", Msg: fmt.Sprintf("cannot check whether %s is in the groups this profile may change: %v", v, err)}
		}
		if len(out.Value) > 0 {
			return nil
		}
	}
	return &Error{Code: "policyScope", Msg: fmt.Sprintf("%s is outside the groups this connection profile may change", v)}
}
