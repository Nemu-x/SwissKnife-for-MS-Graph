package services

import (
	"fmt"
	"net/url"
	"time"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/ops"
)

// CompromisedRequest is the "this account is compromised" response: lock the
// attacker out, then remove what attackers leave behind in the mailbox.
// Blocking sign-in and revoking sessions always run; the rest is optional.
type CompromisedRequest struct {
	Upn             string `json:"upn"`
	Confirm         string `json:"confirm"`
	ResetPassword   bool   `json:"resetPassword"`
	ResetMfa        bool   `json:"resetMfa"`
	ClearForwarding bool   `json:"clearForwarding"`
	DisableRules    bool   `json:"disableRules"`
}

// SignInRow is one recent sign-in, for the operator to judge what happened.
type SignInRow struct {
	When     string `json:"when"`
	App      string `json:"app"`
	IP       string `json:"ip"`
	Location string `json:"location"`
	Result   string `json:"result"`
}

// CompromisedResult is the playbook outcome plus what the operator needs
// next. TempPassword is shown once and never journaled or audited.
type CompromisedResult struct {
	PlaybookResult
	TempPassword string      `json:"tempPassword,omitempty"`
	SignIns      []SignInRow `json:"signIns"`
}

// Compromised runs the compromised-account response.
func (p *PlaybookService) Compromised(req CompromisedRequest) (*CompromisedResult, error) {
	if err := p.s.GuardDestructiveChecked(req.Upn, req.Confirm); err != nil {
		return nil, err
	}
	if err := targetInScope(p.s, engine.FieldUser, req.Upn); err != nil {
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
	r := &runner{op: op, kind: "compromised", ok: true, journal: p.s.Journal}
	if r.journal != nil {
		r.journal.Begin(op.ID, map[string]any{"kind": "playbook", "playbook": "compromised", "target": req.Upn})
		defer func() {
			r.journal.End(op.ID, map[string]any{"ok": r.ok, "canceled": r.canceled, "steps": len(r.steps)})
		}()
	}
	eng := EngineFor(p.s)
	exec := func(id string, in engine.Inputs) error {
		_, err := eng.Execute(op.Ctx, id, in)
		return err
	}
	res := &CompromisedResult{SignIns: []SignInRow{}}

	// Lock the attacker out first: every later step is cleanup.
	r.do("Block sign-in", req.Upn, func() error {
		return exec("user.signIn", engine.Inputs{"user": req.Upn, "state": "blocked"})
	})
	r.do("Revoke sessions", req.Upn, func() error {
		return exec("user.revokeSessions", engine.Inputs{"user": req.Upn})
	})
	if req.ResetPassword && !r.stop() {
		r.do("Reset password", req.Upn, func() error {
			pw, err := actions.TempPassword()
			if err != nil {
				return err
			}
			err = c.Patch(op.Ctx, "/users/"+url.PathEscape(req.Upn), map[string]any{
				"passwordProfile": map[string]any{"password": pw, "forceChangePasswordNextSignIn": true},
			}, nil)
			if err == nil {
				res.TempPassword = pw
			}
			p.s.Record("users.resetPassword", req.Upn, "forceChange=true via=compromised", err)
			return err
		})
	}
	if req.ResetMfa && !r.stop() {
		r.doD("Reset MFA", req.Upn, func() (string, error) {
			out, err := NewAuthMethodsService(p.s).resetMFA(op.Ctx, req.Upn)
			if err != nil {
				return "", err
			}
			detail := fmt.Sprintf("%v method(s) removed", out["removed"])
			// A method left behind is a way back in for the attacker.
			if failures, _ := out["failures"].(map[string]string); len(failures) > 0 {
				for kind, msg := range failures {
					return detail, fmt.Errorf("%d method(s) could not be removed (e.g. %s: %s)", len(failures), kind, msg)
				}
			}
			return detail, nil
		})
	}
	if req.ClearForwarding && !r.stop() {
		r.do("Clear mailbox forwarding", req.Upn, func() error {
			return exec("mailbox.forwarding", engine.Inputs{"mailbox": req.Upn, "op": "clear"})
		})
	}
	if req.DisableRules && !r.stop() {
		r.doD("Disable suspicious inbox rules", req.Upn, func() (string, error) {
			found, err := eng.Run(op.Ctx, "mail.ruleAudit", engine.Inputs{"user": req.Upn})
			if err != nil {
				return "", err
			}
			// One failing rule must not leave the others on: try each, stop
			// only on cancel, and record what was done either way.
			disabled, failed := 0, 0
			var firstErr error
			for _, row := range found.Rows {
				if op.Canceled() {
					break
				}
				if row["ruleId"] == "" || row["enabled"] != "yes" {
					continue
				}
				path := "/users/" + url.PathEscape(req.Upn) + "/mailFolders/inbox/messageRules/" + url.PathEscape(row["ruleId"])
				if err := c.Patch(op.Ctx, path, map[string]any{"isEnabled": false}, nil); err != nil {
					failed++
					if firstErr == nil {
						firstErr = err
					}
					continue
				}
				disabled++
			}
			p.s.Record("mail.disableRules", req.Upn, fmt.Sprintf("disabled=%d failed=%d", disabled, failed), firstErr)
			detail := fmt.Sprintf("%d rule(s) disabled", disabled)
			if firstErr != nil {
				return detail, fmt.Errorf("%d rule(s) could not be disabled: %w", failed, firstErr)
			}
			return detail, nil
		})
	}
	if !r.stop() {
		// Best effort: sign-in logs need Entra ID P1/P2 and AuditLog.Read.All;
		// a failure here does not fail the response.
		var logs struct {
			Value []struct {
				Created  time.Time `json:"createdDateTime"`
				App      string    `json:"appDisplayName"`
				IP       string    `json:"ipAddress"`
				Location struct {
					City    string `json:"city"`
					Country string `json:"countryOrRegion"`
				} `json:"location"`
				Status struct {
					ErrorCode int `json:"errorCode"`
				} `json:"status"`
			} `json:"value"`
		}
		q := url.Values{"$filter": {"userPrincipalName eq '" + escapeODataLiteral(req.Upn) + "'"}, "$top": {"20"}}
		if c.Get(op.Ctx, "/auditLogs/signIns", q, &logs) == nil {
			for _, l := range logs.Value {
				result := "ok"
				if l.Status.ErrorCode != 0 {
					result = fmt.Sprintf("error %d", l.Status.ErrorCode)
				}
				res.SignIns = append(res.SignIns, SignInRow{When: l.Created.Format(time.RFC3339), App: l.App, IP: l.IP,
					Location: joinNonEmpty(l.Location.City, l.Location.Country), Result: result})
			}
		}
	}

	p.recordSummary("summary.compromised", req.Upn, r)
	res.PlaybookResult = *r.result()
	return res, nil
}

func joinNonEmpty(parts ...string) string {
	out := ""
	for _, p := range parts {
		if p == "" {
			continue
		}
		if out != "" {
			out += ", "
		}
		out += p
	}
	return out
}
