package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// PlaybookService runs multi-step composite operations (onboard / offboard) and
// returns a per-step report so the operator sees exactly what happened.
type PlaybookService struct {
	s *session.Session
}

func NewPlaybookService(s *session.Session) *PlaybookService { return &PlaybookService{s: s} }

// Cancel cancels the live playbook operation: its context aborts the in-flight
// step, and a running OneDrive backup (a child operation) cancels with it.
func (p *PlaybookService) Cancel() { p.s.Ops.CancelKind(ops.KindPlaybook) }

// Step is one action within a playbook run. Name/Detail are English fallbacks;
// NameKey/DetailKey (+Params) are stable i18n keys the frontend translates, so
// reports render fully in the UI language (backend-i18n capability).
type Step struct {
	Name      string         `json:"name"`
	NameKey   string         `json:"nameKey,omitempty"`
	OK        bool           `json:"ok"`
	Detail    string         `json:"detail,omitempty"`
	DetailKey string         `json:"detailKey,omitempty"`
	Params    map[string]any `json:"params,omitempty"`
	Error     string         `json:"error,omitempty"`
	ErrorCode string         `json:"errorCode,omitempty"` // Graph error code, when the step failed on a Graph call
	Hint      string         `json:"hint,omitempty"`      // missing Graph permission, when derivable (403)
}

// stepKeys maps step names to stable i18n keys; the English name stays in the
// payload as the fallback for unknown keys.
var stepKeys = map[string]string{
	"Create user":               "steps.createUser",
	"Assign licenses":           "steps.assignLicenses",
	"Add to group":              "steps.addToGroup",
	"Add to team":               "steps.addToTeam",
	"Add to channel":            "steps.addToChannel",
	"Block sign-in":             "steps.blockSignIn",
	"Revoke sessions":           "steps.revokeSessions",
	"Set auto-reply (OOF)":      "steps.oof",
	"Forward mail (inbox rule)": "steps.forward",
	"Hide from address lists":   "steps.hideFromGal",
	"Share calendar (read)":     "steps.shareCalendar",
	"Scan OneDrive":             "steps.scanOneDrive",
	"Backup OneDrive":           "steps.backupOneDrive",
	"Remove from groups":        "steps.removeFromGroups",
	"Remove from group":         "steps.removeFromGroup",
	"Remove licenses":           "steps.removeLicenses",
	"Delete user":               "steps.deleteUser",
}

// PlaybookResult is the outcome of a playbook run.
type PlaybookResult struct {
	OK       bool   `json:"ok"`
	Canceled bool   `json:"canceled"`
	Steps    []Step `json:"steps"`
}

type runner struct {
	op       *ops.Operation
	kind     string // "onboard" | "offboard" — lets the UI route events
	journal  *journal.Log
	steps    []Step
	ok       bool
	canceled bool
	// pending detail translation set by the running step's fn (setDetail).
	pendingDetailKey string
	pendingParams    map[string]any
}

// setDetail lets a step body attach a translatable detail (key + params) to
// the step it is running in, alongside the English fallback it returns.
func (r *runner) setDetail(key string, params map[string]any) {
	r.pendingDetailKey, r.pendingParams = key, params
}

var errPlaybookCanceled = fmt.Errorf("playbook canceled")

// emitStep streams one step lifecycle event so the UI can render live progress
// instead of waiting for the whole playbook to finish, and journals completed
// steps so the run survives an app restart.
func (r *runner) emitStep(payload map[string]any) {
	payload["kind"] = r.kind
	emitOp(r.op.Ctx, r.op, "playbook:step", payload)
	if r.journal != nil && payload["status"] == "done" {
		r.journal.Event(r.op.ID, "step", payload)
	}
}

func (r *runner) do(name, detail string, fn func() error) error {
	return r.doD(name, detail, func() (string, error) { return "", fn() })
}

// stop reports whether a cancel was requested; once true, remaining steps are
// skipped without being reported.
func (r *runner) stop() bool {
	if r.op.Canceled() {
		r.canceled = true
	}
	return r.canceled
}

// doD is do with a detail returned by the step itself (e.g. a scanned size),
// which replaces the static detail when non-empty.
func (r *runner) doD(name, detail string, fn func() (string, error)) error {
	if r.stop() {
		return errPlaybookCanceled
	}
	r.pendingDetailKey, r.pendingParams = "", nil
	r.emitStep(map[string]any{"status": "running", "index": len(r.steps), "name": name, "nameKey": stepKeys[name], "detail": detail})
	d, err := fn()
	if d != "" {
		detail = d
	}
	st := Step{Name: name, NameKey: stepKeys[name], OK: err == nil, Detail: detail,
		DetailKey: r.pendingDetailKey, Params: r.pendingParams}
	if err != nil {
		st.Error = err.Error()
		r.ok = false
		// Surface Graph error structure so the UI can show actionable hints
		// ("grant Mail.ReadWrite in Entra") instead of a bare 403 string.
		var ge *graphapi.GraphError
		if errors.As(err, &ge) {
			st.Error = ge.Message
			st.ErrorCode = ge.Code
			if ge.StatusCode == 403 {
				st.Hint = permissionHint(ge.Path)
			}
		}
	}
	r.steps = append(r.steps, st)
	done := map[string]any{"status": "done", "index": len(r.steps) - 1, "name": name, "nameKey": st.NameKey, "detail": detail, "ok": st.OK}
	if st.DetailKey != "" {
		done["detailKey"] = st.DetailKey
		done["params"] = st.Params
	}
	if st.Error != "" {
		done["error"] = st.Error
	}
	if st.ErrorCode != "" {
		done["errorCode"] = st.ErrorCode
	}
	if st.Hint != "" {
		done["hint"] = st.Hint
	}
	r.emitStep(done)
	return err
}

func (r *runner) result() *PlaybookResult {
	return &PlaybookResult{OK: r.ok, Canceled: r.canceled, Steps: r.steps}
}

// ChannelRef targets a specific channel within a team.
type ChannelRef struct {
	TeamID    string `json:"teamId"`
	ChannelID string `json:"channelId"`
}

// OnboardRequest describes a new-hire onboarding run.
type OnboardRequest struct {
	DisplayName   string       `json:"displayName"`
	Upn           string       `json:"upn"`
	MailNickname  string       `json:"mailNickname"`
	Password      string       `json:"password"`
	UsageLocation string       `json:"usageLocation"`
	SkuIDs        []string     `json:"skuIds"`
	GroupIDs      []string     `json:"groupIds"`
	TeamIDs       []string     `json:"teamIds"`
	ChannelRefs   []ChannelRef `json:"channelRefs"`
}

// Onboard creates the user, then (best-effort) sets usage location, assigns
// licenses, and adds the user to groups and teams.
func (p *PlaybookService) Onboard(req OnboardRequest) (*PlaybookResult, error) {
	if err := p.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := p.s.Client()
	if err != nil {
		return nil, err
	}
	op, err := p.s.Ops.Start(p.s.Ctx(), ops.KindPlaybook)
	if err != nil {
		return nil, err
	}
	defer p.s.Ops.Finish(op)
	emitOp(p.s.Ctx(), op, "op:start", map[string]any{"target": req.Upn})
	r := &runner{op: op, kind: "onboard", ok: true, journal: p.s.Journal}
	if r.journal != nil {
		r.journal.Begin(op.ID, map[string]any{"kind": "playbook", "playbook": "onboard", "target": req.Upn})
		defer func() { r.journal.End(op.ID, map[string]any{"ok": r.ok, "canceled": r.canceled, "steps": len(r.steps)}) }()
	}

	// Step 1: create user (blocks the rest if it fails).
	body := map[string]any{
		"accountEnabled":    true,
		"displayName":       req.DisplayName,
		"mailNickname":      req.MailNickname,
		"userPrincipalName": req.Upn,
		"passwordProfile":   map[string]any{"password": req.Password, "forceChangePasswordNextSignIn": true},
	}
	if req.UsageLocation != "" {
		body["usageLocation"] = req.UsageLocation
	}
	createErr := r.do("Create user", req.Upn, func() error {
		return c.Post(op.Ctx, "/users", body, nil)
	})
	if createErr != nil {
		p.s.Record("playbook.onboard", req.Upn, "create failed", createErr)
		return r.result(), nil
	}

	// resolve the new user's object id for group $ref bindings
	var created struct {
		ID string `json:"id"`
	}
	_ = c.Get(op.Ctx, "/users/"+url.PathEscape(req.Upn), url.Values{"$select": {"id"}}, &created)

	if len(req.SkuIDs) > 0 {
		add := make([]map[string]any, 0, len(req.SkuIDs))
		for _, id := range req.SkuIDs {
			add = append(add, map[string]any{"skuId": id})
		}
		r.do("Assign licenses", itoa(len(req.SkuIDs))+" sku(s)", func() error {
			return c.Post(op.Ctx, "/users/"+url.PathEscape(req.Upn)+"/assignLicense",
				map[string]any{"addLicenses": add, "removeLicenses": []string{}}, nil)
		})
	}

	for _, gid := range req.GroupIDs {
		r.do("Add to group", gid, func() error {
			ref := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + created.ID}
			return c.Post(op.Ctx, "/groups/"+url.PathEscape(gid)+"/members/$ref", ref, nil)
		})
	}

	for _, tid := range req.TeamIDs {
		r.do("Add to team", tid, func() error {
			return c.Post(op.Ctx, "/teams/"+url.PathEscape(tid)+"/members", conversationMember(req.Upn, false), nil)
		})
	}

	for _, cr := range req.ChannelRefs {
		cr := cr
		r.do("Add to channel", cr.TeamID+"/"+cr.ChannelID, func() error {
			return c.Post(op.Ctx,
				"/teams/"+url.PathEscape(cr.TeamID)+"/channels/"+url.PathEscape(cr.ChannelID)+"/members",
				conversationMember(req.Upn, false), nil)
		})
	}

	p.s.Record("playbook.onboard", req.Upn, "steps="+itoa(len(r.steps)), nil)
	p.recordSummary("summary.onboard", req.Upn, r)
	return r.result(), nil
}

// OffboardRequest describes an employee offboarding run.
type OffboardRequest struct {
	Upn               string `json:"upn"`
	Confirm           string `json:"confirm"`
	Block             bool   `json:"block"`
	RevokeSessions    bool   `json:"revokeSessions"`
	Oof               bool   `json:"oof"`
	OofMessage        string `json:"oofMessage"`
	ForwardTo         string `json:"forwardTo"`
	HideFromGal       bool   `json:"hideFromGal"`
	CalendarTo        string `json:"calendarTo"`
	RemoveFromGroups  bool   `json:"removeFromGroups"`
	RemoveAllLicenses bool   `json:"removeAllLicenses"`
	BackupToUser      string `json:"backupToUser"`
	BackupFolder      string `json:"backupFolder"`
	Delete            bool   `json:"delete"`
}

// Offboard runs the offboarding sequence. Destructive: requires typed confirm
// on the UPN. Steps are best-effort and reported individually.
func (p *PlaybookService) Offboard(req OffboardRequest) (*PlaybookResult, error) {
	if err := p.s.GuardDestructive(req.Upn, req.Confirm); err != nil {
		return nil, err
	}
	c, err := p.s.Client()
	if err != nil {
		return nil, err
	}
	op, err := p.s.Ops.Start(p.s.Ctx(), ops.KindPlaybook)
	if err != nil {
		return nil, err
	}
	defer p.s.Ops.Finish(op)
	emitOp(p.s.Ctx(), op, "op:start", map[string]any{"target": req.Upn})
	r := &runner{op: op, kind: "offboard", ok: true, journal: p.s.Journal}
	if r.journal != nil {
		r.journal.Begin(op.ID, map[string]any{"kind": "playbook", "playbook": "offboard", "target": req.Upn})
		defer func() { r.journal.End(op.ID, map[string]any{"ok": r.ok, "canceled": r.canceled, "steps": len(r.steps)}) }()
	}
	u := url.PathEscape(req.Upn)

	if req.Block {
		r.do("Block sign-in", req.Upn, func() error {
			return c.Patch(op.Ctx, "/users/"+u, map[string]any{"accountEnabled": false}, nil)
		})
	}
	if req.RevokeSessions {
		r.do("Revoke sessions", req.Upn, func() error {
			return c.Post(op.Ctx, "/users/"+u+"/revokeSignInSessions", map[string]any{}, nil)
		})
	}
	if req.Oof {
		r.do("Set auto-reply (OOF)", req.Upn, func() error {
			msg := req.OofMessage
			if msg == "" {
				msg = "This employee is no longer with the organization. Your message will not be forwarded automatically."
			}
			return c.Patch(op.Ctx, "/users/"+u+"/mailboxSettings", map[string]any{
				"automaticRepliesSetting": map[string]any{
					"status":               "alwaysEnabled",
					"internalReplyMessage": msg,
					"externalReplyMessage": msg,
				},
			}, nil)
		})
	}
	if req.ForwardTo != "" {
		// Server-side inbox rule; unlike Exchange mailbox forwarding this is
		// available through Graph (Mail.ReadWrite) and survives sign-in block.
		r.do("Forward mail (inbox rule)", req.Upn+" → "+req.ForwardTo, func() error {
			return c.Post(op.Ctx, "/users/"+u+"/mailFolders/inbox/messageRules", map[string]any{
				"displayName": "Offboarding: forward to " + req.ForwardTo,
				"sequence":    1,
				"isEnabled":   true,
				"actions": map[string]any{
					"forwardTo": []any{
						map[string]any{"emailAddress": map[string]any{"address": req.ForwardTo}},
					},
					"stopProcessingRules": false,
				},
			}, nil)
		})
	}
	if req.HideFromGal {
		r.do("Hide from address lists", req.Upn, func() error {
			return c.Patch(op.Ctx, "/users/"+u, map[string]any{"showInAddressList": false}, nil)
		})
	}
	if req.CalendarTo != "" {
		r.do("Share calendar (read)", req.Upn+" → "+req.CalendarTo, func() error {
			return c.Post(op.Ctx, "/users/"+u+"/calendar/calendarPermissions", map[string]any{
				"emailAddress": map[string]any{"address": req.CalendarTo},
				"role":         "read",
			}, nil)
		})
	}
	if req.BackupToUser != "" {
		drive := NewDriveService(p.s)
		// Scan first so the operator sees the transfer volume up front, and so
		// the copy can report overall percent progress against a known total.
		var prev *CopyPreview
		r.doD("Scan OneDrive", req.Upn, func() (string, error) {
			pv, e := drive.OffboardingPreview(req.Upn)
			if e != nil {
				return "", e
			}
			prev = pv
			r.setDetail("stepDetails.scanned", map[string]any{"files": pv.Files, "bytes": pv.TotalBytes})
			return itoa(pv.Files) + " files · " + humanSize(pv.TotalBytes), nil
		})
		r.doD("Backup OneDrive", req.Upn+" → "+req.BackupToUser, func() (string, error) {
			folder := req.BackupFolder
			if folder == "" {
				folder = req.Upn
			}
			// The copy is a child operation: cancelling the playbook cancels it
			// through context parentage.
			res, e := drive.copyBetweenUsersCtx(op.Ctx, req.Upn, req.BackupToUser, folder, false, prev)
			if e != nil {
				return "", e
			}
			params := map[string]any{"copied": len(res.Copied), "skipped": len(res.Skipped), "canceled": res.Canceled}
			detail := itoa(len(res.Copied)) + " item(s) copied"
			if prev != nil {
				params["bytes"] = prev.TotalBytes
				detail += " · " + humanSize(prev.TotalBytes)
			}
			if len(res.Skipped) > 0 {
				detail += " · " + itoa(len(res.Skipped)) + " skipped"
			}
			if res.Canceled {
				detail += " · canceled"
			}
			r.setDetail("stepDetails.backup", params)
			if len(res.Failed) > 0 {
				return detail, fmt.Errorf("%d file(s) failed — see the OneDrive transfer log", len(res.Failed))
			}
			return detail, nil
		})
	}
	if req.RemoveFromGroups && !r.stop() {
		// One report step per group so the operator sees exactly what happened.
		// Dynamic-membership and Exchange-managed (distribution) groups fail
		// individually with the Graph error; the rest still get removed.
		var idResp struct {
			ID string `json:"id"`
		}
		if err := c.Get(op.Ctx, "/users/"+u, url.Values{"$select": {"id"}}, &idResp); err != nil {
			r.do("Remove from groups", req.Upn, func() error { return err })
		} else if items, err := c.ListAll(op.Ctx, "/users/"+u+"/memberOf", url.Values{"$select": {"id,displayName"}}, 0); err != nil {
			r.do("Remove from groups", req.Upn, func() error { return err })
		} else {
			for _, raw := range items {
				var g struct {
					Type        string `json:"@odata.type"`
					ID          string `json:"id"`
					DisplayName string `json:"displayName"`
				}
				if json.Unmarshal(raw, &g) != nil || g.Type != "#microsoft.graph.group" {
					continue
				}
				gid := g.ID
				name := g.DisplayName
				if name == "" {
					name = gid
				}
				r.do("Remove from group", name, func() error {
					return c.Delete(op.Ctx, "/groups/"+url.PathEscape(gid)+"/members/"+url.PathEscape(idResp.ID)+"/$ref")
				})
			}
		}
	}
	if req.RemoveAllLicenses {
		r.do("Remove licenses", req.Upn, func() error {
			var lic struct {
				Value []struct {
					SkuID string `json:"skuId"`
				} `json:"value"`
			}
			if e := c.Get(op.Ctx, "/users/"+u+"/licenseDetails", nil, &lic); e != nil {
				return e
			}
			remove := make([]string, 0, len(lic.Value))
			for _, l := range lic.Value {
				remove = append(remove, l.SkuID)
			}
			if len(remove) == 0 {
				return nil
			}
			return c.Post(op.Ctx, "/users/"+u+"/assignLicense",
				map[string]any{"addLicenses": []any{}, "removeLicenses": remove}, nil)
		})
	}
	if req.Delete {
		r.do("Delete user", req.Upn, func() error {
			return c.Delete(op.Ctx, "/users/"+u)
		})
	}

	p.s.Record("playbook.offboard", req.Upn, "steps="+itoa(len(r.steps)), nil)
	p.recordSummary("summary.offboard", req.Upn, r)
	return r.result(), nil
}

// recordSummary writes one human-readable (translatable key+params) audit
// entry per playbook run, so Activity reads "offboarded X: N steps, M failed".
func (p *PlaybookService) recordSummary(key, upn string, r *runner) {
	failed := 0
	for _, s := range r.steps {
		if !s.OK {
			failed++
		}
	}
	detail, err := json.Marshal(map[string]any{
		"key":    key,
		"params": map[string]any{"upn": upn, "steps": len(r.steps), "failed": failed, "canceled": r.canceled},
	})
	if err != nil {
		return
	}
	p.s.Record("playbook.summary", upn, string(detail), nil)
}

var _ = json.Marshal
