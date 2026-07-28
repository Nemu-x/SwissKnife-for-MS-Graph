package services

import (
	"context"
	"encoding/json"
	"errors"
	"net/url"
	"strings"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// MirrorService answers the everyday ticket "give Ivanov the same access Petrov
// has": it reads both access profiles, diffs them, and copies what the target is
// missing. Additive only — it never removes anything the target already has, so
// a mirror run can only widen access, never silently narrow it.
type MirrorService struct {
	s *session.Session
}

func NewMirrorService(s *session.Session) *MirrorService { return &MirrorService{s: s} }

// Cancel aborts a running mirror comparison or copy.
func (m *MirrorService) Cancel() { m.s.Ops.CancelKind(ops.KindMirror) }

// Access item kinds, also used as the selection keys of a copy request.
const (
	KindGroup   = "group"
	KindRole    = "role"
	KindTeam    = "team"
	KindChannel = "channel"
	KindLicense = "license"
)

// AccessRow is one piece of access, either side of the comparison. Status says
// what the copy would do with it; ReasonKey explains a Copyable=false.
type AccessRow struct {
	Kind      string `json:"kind"`
	ID        string `json:"id"`
	Name      string `json:"name"`
	TeamID    string `json:"teamId,omitempty"`
	TeamName  string `json:"teamName,omitempty"`
	Sub       string `json:"sub,omitempty"`    // membershipType, SKU part number…
	Status    string `json:"status"`           // missing | both | targetOnly
	Copyable  bool   `json:"copyable"`
	ReasonKey string `json:"reasonKey,omitempty"`
}

// MirrorRequest is a copy run: what to take from Source and give to Target.
type MirrorRequest struct {
	Source  string   `json:"source"`
	Target  string   `json:"target"`
	Kinds   []string `json:"kinds"` // subset of the Kind* constants
	Confirm string   `json:"confirm"`
}

type mirrorUser struct {
	id  string
	upn string
}

// resolveUserID turns a UPN into an object id. Deep paths built from a UPN break
// Graph's OData parser when the UPN holds an apostrophe, a leading '$' or the
// guest '#EXT#' marker, so callers that build /users/{x}/... address by GUID.
// A value that already looks like a GUID is passed through unchanged.
func resolveUserID(ctx context.Context, c *graphapi.Client, upn string) (string, error) {
	if looksLikeGUID(upn) {
		return upn, nil
	}
	var out struct {
		ID string `json:"id"`
	}
	if err := c.Get(ctx, "/users/"+url.PathEscape(upn), url.Values{"$select": {"id"}}, &out); err != nil {
		return "", err
	}
	if out.ID == "" {
		return "", errors.New("user not found: " + upn)
	}
	return out.ID, nil
}

func looksLikeGUID(s string) bool {
	if len(s) != 36 {
		return false
	}
	for i, ch := range s {
		switch i {
		case 8, 13, 18, 23:
			if ch != '-' {
				return false
			}
		default:
			isHex := (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F')
			if !isHex {
				return false
			}
		}
	}
	return true
}

func (m *MirrorService) resolveUser(upn string) (mirrorUser, error) {
	c, err := m.s.Client()
	if err != nil {
		return mirrorUser{}, err
	}
	var out struct {
		ID  string `json:"id"`
		Upn string `json:"userPrincipalName"`
	}
	if err := c.Get(m.s.Ctx(), "/users/"+url.PathEscape(upn), url.Values{"$select": {"id,userPrincipalName"}}, &out); err != nil {
		return mirrorUser{}, err
	}
	if out.ID == "" {
		return mirrorUser{}, errors.New("user not found: " + upn)
	}
	return mirrorUser{id: out.ID, upn: out.Upn}, nil
}

// groupCopyable decides whether Graph can add a member to this group at all.
// Dynamic membership is computed from a rule, and Exchange-managed groups
// (distribution lists, mail-enabled security) only change in Exchange.
func groupCopyable(groupTypes []string, mailEnabled, securityEnabled bool) (bool, string) {
	unified := false
	for _, gt := range groupTypes {
		if strings.EqualFold(gt, "Unified") {
			unified = true
		}
		if strings.Contains(strings.ToLower(gt), "dynamic") {
			return false, "reasons.dynamicGroup"
		}
	}
	if unified {
		return true, ""
	}
	if mailEnabled {
		return false, "reasons.exchangeGroup"
	}
	if securityEnabled {
		return true, ""
	}
	return true, ""
}

// profile collects everything one user has: groups, admin roles, teams, private
// and shared channels, licenses. Progress is streamed because the channel walk
// is one Graph call per private channel and can run for a while on a person who
// sits in twenty teams.
func (m *MirrorService) profile(u mirrorUser, op *ops.Operation, side string) ([]AccessRow, error) {
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	ctx := m.s.Ctx()
	base := "/users/" + url.PathEscape(u.id)
	rows := []AccessRow{}
	progress := func(what, name string, done, total int) {
		emitOp(ctx, op, "mirror:progress", map[string]any{
			"side": side, "what": what, "name": name, "done": done, "total": total,
		})
	}

	// Groups and directory roles arrive in one collection, told apart by type.
	progress("groups", "", 0, 0)
	memberOf, err := c.ListAll(ctx, base+"/memberOf",
		url.Values{"$select": {"id,displayName,groupTypes,mailEnabled,securityEnabled"}}, 0)
	if err != nil {
		return nil, err
	}
	for _, raw := range memberOf {
		var g struct {
			Type            string   `json:"@odata.type"`
			ID              string   `json:"id"`
			DisplayName     string   `json:"displayName"`
			GroupTypes      []string `json:"groupTypes"`
			MailEnabled     bool     `json:"mailEnabled"`
			SecurityEnabled bool     `json:"securityEnabled"`
		}
		if json.Unmarshal(raw, &g) != nil || g.ID == "" {
			continue
		}
		if strings.Contains(g.Type, "directoryRole") {
			rows = append(rows, AccessRow{Kind: KindRole, ID: g.ID, Name: g.DisplayName, Copyable: true})
			continue
		}
		ok, reason := groupCopyable(g.GroupTypes, g.MailEnabled, g.SecurityEnabled)
		rows = append(rows, AccessRow{Kind: KindGroup, ID: g.ID, Name: g.DisplayName, Copyable: ok, ReasonKey: reason})
	}

	// Teams, then the channels of each team that have their own membership.
	progress("teams", "", 0, 0)
	teams, err := c.ListAll(ctx, base+"/joinedTeams", url.Values{"$select": {"id,displayName"}}, 0)
	if err != nil {
		return nil, err
	}
	for i, raw := range teams {
		var tm struct {
			ID          string `json:"id"`
			DisplayName string `json:"displayName"`
		}
		if json.Unmarshal(raw, &tm) != nil || tm.ID == "" {
			continue
		}
		rows = append(rows, AccessRow{Kind: KindTeam, ID: tm.ID, Name: tm.DisplayName, Copyable: true})

		if ctx.Err() != nil {
			return rows, ctx.Err()
		}
		progress("channels", tm.DisplayName, i+1, len(teams))
		chans, err := c.ListAll(ctx, "/teams/"+url.PathEscape(tm.ID)+"/allChannels",
			url.Values{"$select": {"id,displayName,membershipType"}}, 0)
		if err != nil {
			continue // a team we cannot read channels for must not kill the whole scan
		}
		for _, cr := range chans {
			var ch struct {
				ID             string `json:"id"`
				DisplayName    string `json:"displayName"`
				MembershipType string `json:"membershipType"`
			}
			if json.Unmarshal(cr, &ch) != nil || ch.ID == "" || ch.MembershipType == "" ||
				strings.EqualFold(ch.MembershipType, "standard") {
				continue // standard channels inherit the team's membership
			}
			member, err := m.inChannel(tm.ID, ch.ID, u)
			if err != nil || !member {
				continue
			}
			rows = append(rows, AccessRow{
				Kind: KindChannel, ID: ch.ID, Name: ch.DisplayName,
				TeamID: tm.ID, TeamName: tm.DisplayName, Sub: ch.MembershipType, Copyable: true,
			})
		}
	}

	progress("licenses", "", 0, 0)
	licenses, err := c.ListAll(ctx, base+"/licenseDetails", url.Values{"$select": {"skuId,skuPartNumber"}}, 0)
	if err != nil {
		return nil, err
	}
	for _, raw := range licenses {
		var l struct {
			SkuID         string `json:"skuId"`
			SkuPartNumber string `json:"skuPartNumber"`
		}
		if json.Unmarshal(raw, &l) != nil || l.SkuID == "" {
			continue
		}
		rows = append(rows, AccessRow{Kind: KindLicense, ID: l.SkuID, Name: l.SkuPartNumber, Copyable: true})
	}
	return rows, nil
}

// inChannel reports whether the user holds a membership in this channel.
func (m *MirrorService) inChannel(teamID, channelID string, u mirrorUser) (bool, error) {
	c, err := m.s.Client()
	if err != nil {
		return false, err
	}
	members, err := c.ListAll(m.s.Ctx(),
		"/teams/"+url.PathEscape(teamID)+"/channels/"+url.PathEscape(channelID)+"/members", nil, 0)
	if err != nil {
		return false, err
	}
	for _, raw := range members {
		var mem struct {
			UserID string `json:"userId"`
			Email  string `json:"email"`
		}
		if json.Unmarshal(raw, &mem) != nil {
			continue
		}
		if mem.UserID == u.id || (mem.Email != "" && strings.EqualFold(mem.Email, u.upn)) {
			return true, nil
		}
	}
	return false, nil
}

func rowKey(r AccessRow) string { return r.Kind + "|" + r.TeamID + "|" + r.ID }

// Compare is the read-only half: what the source has, what the target has, and
// what a copy would add. Rows come back flat so the results view can show them.
func (m *MirrorService) Compare(sourceUpn, targetUpn string) ([]AccessRow, error) {
	src, err := m.resolveUser(sourceUpn)
	if err != nil {
		return nil, err
	}
	tgt, err := m.resolveUser(targetUpn)
	if err != nil {
		return nil, err
	}
	op, err := m.s.Ops.Start(m.s.Ctx(), ops.KindMirror)
	if err != nil {
		return nil, err
	}
	defer m.s.Ops.Finish(op)
	emitOp(m.s.Ctx(), op, "op:start", map[string]any{"target": targetUpn})

	srcRows, err := m.profile(src, op, "source")
	if err != nil {
		return nil, err
	}
	tgtRows, err := m.profile(tgt, op, "target")
	if err != nil {
		return nil, err
	}
	have := map[string]bool{}
	for _, r := range tgtRows {
		have[rowKey(r)] = true
	}
	out := make([]AccessRow, 0, len(srcRows)+len(tgtRows))
	for _, r := range srcRows {
		if have[rowKey(r)] {
			r.Status = "both"
			r.Copyable = false
		} else {
			r.Status = "missing"
		}
		out = append(out, r)
	}
	// What only the target has is not copied, but the operator should see it.
	srcKeys := map[string]bool{}
	for _, r := range srcRows {
		srcKeys[rowKey(r)] = true
	}
	for _, r := range tgtRows {
		if !srcKeys[rowKey(r)] {
			r.Status = "targetOnly"
			r.Copyable = false
			out = append(out, r)
		}
	}
	m.s.Record("mirror.compare", targetUpn, "source="+sourceUpn+" rows="+itoa(len(out)), nil)
	return out, nil
}

// Copy grants the target everything the source has and the target lacks, for the
// selected kinds. Every item is its own step, so a partial failure is visible
// and the rest still runs.
func (m *MirrorService) Copy(req MirrorRequest) (*PlaybookResult, error) {
	if err := m.s.GuardWrite(); err != nil {
		return nil, err
	}
	// Granting somebody else's access wholesale deserves the typed confirmation.
	if err := m.s.GuardDestructive(req.Target, req.Confirm); err != nil {
		return nil, err
	}
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	src, err := m.resolveUser(req.Source)
	if err != nil {
		return nil, err
	}
	tgt, err := m.resolveUser(req.Target)
	if err != nil {
		return nil, err
	}
	want := map[string]bool{}
	for _, k := range req.Kinds {
		want[k] = true
	}

	rows, err := m.Compare(req.Source, req.Target)
	if err != nil {
		return nil, err
	}

	op, err := m.s.Ops.Start(m.s.Ctx(), ops.KindMirror)
	if err != nil {
		return nil, err
	}
	defer m.s.Ops.Finish(op)
	emitOp(m.s.Ctx(), op, "op:start", map[string]any{"target": req.Target})
	r := &runner{op: op, kind: "mirror", ok: true, journal: m.s.Journal}
	if r.journal != nil {
		r.journal.Begin(op.ID, map[string]any{"kind": "mirror", "source": src.upn, "target": tgt.upn})
		defer func() {
			r.journal.End(op.ID, map[string]any{"ok": r.ok, "canceled": r.canceled, "steps": len(r.steps)})
		}()
	}

	ref := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + tgt.id}
	// Teams before channels: a private channel only accepts existing members of
	// the team, so a channel copy has to bring its team along.
	order := []string{KindLicense, KindGroup, KindRole, KindTeam, KindChannel}
	teamsAdded := map[string]bool{}

	for _, kind := range order {
		if !want[kind] {
			continue
		}
		for _, row := range rows {
			if row.Kind != kind || row.Status != "missing" {
				continue
			}
			row := row
			if !row.Copyable {
				r.steps = append(r.steps, Step{
					Name: skipName(kind), NameKey: stepKeys[skipName(kind)], OK: false,
					Detail: row.Name, Error: "not copyable via Graph", DetailKey: row.ReasonKey,
				})
				r.ok = false
				continue
			}
			switch kind {
			case KindLicense:
				r.do("Assign licenses", row.Name, func() error {
					return c.Post(op.Ctx, "/users/"+url.PathEscape(tgt.id)+"/assignLicense",
						map[string]any{"addLicenses": []map[string]any{{"skuId": row.ID}}, "removeLicenses": []string{}}, nil)
				})
			case KindGroup:
				r.do("Add to group", row.Name, func() error {
					return c.Post(op.Ctx, "/groups/"+url.PathEscape(row.ID)+"/members/$ref", ref, nil)
				})
			case KindRole:
				r.do("Add to role", row.Name, func() error {
					return c.Post(op.Ctx, "/directoryRoles/"+url.PathEscape(row.ID)+"/members/$ref", ref, nil)
				})
			case KindTeam:
				if err := r.do("Add to team", row.Name, func() error {
					return c.Post(op.Ctx, "/teams/"+url.PathEscape(row.ID)+"/members", conversationMember(tgt.upn, false), nil)
				}); err == nil {
					teamsAdded[row.ID] = true
				}
			case KindChannel:
				// The team membership may be missing when only channels were
				// selected — add it first, once per team, and say so.
				if !teamsAdded[row.TeamID] && !m.targetInTeam(rows, row.TeamID) {
					if err := r.do("Add to team", row.TeamName, func() error {
						return c.Post(op.Ctx, "/teams/"+url.PathEscape(row.TeamID)+"/members", conversationMember(tgt.upn, false), nil)
					}); err == nil {
						teamsAdded[row.TeamID] = true
					}
				}
				r.do("Add to channel", row.TeamName+" / "+row.Name, func() error {
					return c.Post(op.Ctx,
						"/teams/"+url.PathEscape(row.TeamID)+"/channels/"+url.PathEscape(row.ID)+"/members",
						conversationMember(tgt.upn, false), nil)
				})
			}
			if r.stop() {
				break
			}
		}
		if r.stop() {
			break
		}
	}

	m.s.Record("mirror.copy", req.Target, "source="+req.Source+" steps="+itoa(len(r.steps)), nil)
	return r.result(), nil
}

// targetInTeam reports whether the target already belongs to the team (the
// comparison rows carry the target's own memberships as "both"/"targetOnly").
func (m *MirrorService) targetInTeam(rows []AccessRow, teamID string) bool {
	for _, r := range rows {
		if r.Kind == KindTeam && r.ID == teamID && (r.Status == "both" || r.Status == "targetOnly") {
			return true
		}
	}
	return false
}

// skipName is the step name used to report an item Graph cannot copy.
func skipName(kind string) string {
	switch kind {
	case KindGroup:
		return "Add to group"
	case KindRole:
		return "Add to role"
	case KindTeam:
		return "Add to team"
	case KindChannel:
		return "Add to channel"
	default:
		return "Assign licenses"
	}
}
