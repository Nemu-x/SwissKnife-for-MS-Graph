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
	if onDirectory(a) {
		return nil // tenant profile limits do not govern the on-prem directory
	}
	pol := e.s.Policy()
	if !dangerAllowed(effectiveDanger(a), pol.MaxDanger) {
		return &Error{Code: "policyDanger", Msg: fmt.Sprintf("this connection profile allows %s actions at most", pol.MaxDanger)}
	}
	if len(pol.AllowedGroups) == 0 {
		return nil
	}
	// A pack script can touch any object whatever its inputs name: under a
	// group scope packs do not run at all.
	if a.Pack != "" {
		return &Error{Code: "policyUnscoped", Msg: "this connection profile may only change members of its groups, and action-pack scripts cannot be checked against them"}
	}
	if a.Danger == Read {
		return nil
	}
	named := 0
	for _, f := range a.Fields {
		v := strings.TrimSpace(in[f.Name])
		if v == "" || (f.Kind != FieldUser && f.Kind != FieldGroup) {
			continue
		}
		named++
		if err := e.inScope(env, f.Kind, v, pol.AllowedGroups); err != nil {
			return err
		}
	}
	// Tenant-wide changes (a purge, a policy restore) name no user or group:
	// a group-scoped profile cannot run them.
	if named == 0 {
		return &Error{Code: "policyUnscoped", Msg: "this connection profile may only change members of its groups, and this action is tenant-wide"}
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
		return &Error{Code: "policyCheckFailed", Msg: "not connected"}
	}
	// checkMemberGroups is called on the object id: some UPNs (quotes, a
	// leading $) do not work as a path segment.
	id := v
	if kind == FieldUser && !looksLikeID(v) {
		var res struct {
			Value []struct {
				ID string `json:"id"`
			} `json:"value"`
		}
		q := url.Values{"$filter": {"userPrincipalName eq '" + strings.ReplaceAll(v, "'", "''") + "'"}, "$select": {"id"}}
		if err := env.Graph.Get(env.Ctx, "/users", q, &res); err != nil {
			return &Error{Code: "policyCheckFailed", Msg: fmt.Sprintf("cannot check whether %s is in the groups this profile may change: %v", v, err)}
		}
		if len(res.Value) != 1 {
			return &Error{Code: "policyCheckFailed", Msg: fmt.Sprintf("user %s not found", v)}
		}
		id = res.Value[0].ID
	}
	for start := 0; start < len(allowed); start += 20 { // checkMemberGroups takes 20 ids
		end := min(start+20, len(allowed))
		var out struct {
			Value []string `json:"value"`
		}
		body := map[string]any{"groupIds": allowed[start:end]}
		if err := env.Graph.Post(env.Ctx, coll+url.PathEscape(id)+"/checkMemberGroups", body, &out); err != nil {
			return &Error{Code: "policyCheckFailed", Msg: fmt.Sprintf("cannot check whether %s is in the groups this profile may change: %v", v, err)}
		}
		if len(out.Value) > 0 {
			return nil
		}
	}
	return &Error{Code: "policyScope", Msg: fmt.Sprintf("%s is outside the groups this connection profile may change", v)}
}

// looksLikeID reports a GUID-shaped object id.
func looksLikeID(v string) bool {
	if len(v) != 36 {
		return false
	}
	for i, r := range v {
		switch {
		case i == 8 || i == 13 || i == 18 || i == 23:
			if r != '-' {
				return false
			}
		case (r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F'):
		default:
			return false
		}
	}
	return true
}

// effectiveDanger is the danger the profile limits judge: a pack's own
// "read" is a claim its script is not held to, so packs count as writes.
func effectiveDanger(a Action) Danger {
	if a.Pack != "" && a.Danger == Read {
		return Write
	}
	return a.Danger
}
