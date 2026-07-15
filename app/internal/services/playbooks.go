package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

// PlaybookService runs multi-step composite operations (onboard / offboard) and
// returns a per-step report so the operator sees exactly what happened.
type PlaybookService struct {
	s *session.Session
}

func NewPlaybookService(s *session.Session) *PlaybookService { return &PlaybookService{s: s} }

// Step is one action within a playbook run.
type Step struct {
	Name   string `json:"name"`
	OK     bool   `json:"ok"`
	Detail string `json:"detail,omitempty"`
	Error  string `json:"error,omitempty"`
}

// PlaybookResult is the outcome of a playbook run.
type PlaybookResult struct {
	OK    bool   `json:"ok"`
	Steps []Step `json:"steps"`
}

type runner struct {
	steps []Step
	ok    bool
}

func (r *runner) do(name, detail string, fn func() error) error {
	err := fn()
	st := Step{Name: name, OK: err == nil, Detail: detail}
	if err != nil {
		st.Error = err.Error()
		r.ok = false
	}
	r.steps = append(r.steps, st)
	return err
}

func (r *runner) result() *PlaybookResult { return &PlaybookResult{OK: r.ok, Steps: r.steps} }

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
	r := &runner{ok: true}

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
		return c.Post(p.s.Ctx(), "/users", body, nil)
	})
	if createErr != nil {
		p.s.Record("playbook.onboard", req.Upn, "create failed", createErr)
		return r.result(), nil
	}

	// resolve the new user's object id for group $ref bindings
	var created struct {
		ID string `json:"id"`
	}
	_ = c.Get(p.s.Ctx(), "/users/"+url.PathEscape(req.Upn), url.Values{"$select": {"id"}}, &created)

	if len(req.SkuIDs) > 0 {
		add := make([]map[string]any, 0, len(req.SkuIDs))
		for _, id := range req.SkuIDs {
			add = append(add, map[string]any{"skuId": id})
		}
		r.do("Assign licenses", itoa(len(req.SkuIDs))+" sku(s)", func() error {
			return c.Post(p.s.Ctx(), "/users/"+url.PathEscape(req.Upn)+"/assignLicense",
				map[string]any{"addLicenses": add, "removeLicenses": []string{}}, nil)
		})
	}

	for _, gid := range req.GroupIDs {
		r.do("Add to group", gid, func() error {
			ref := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + created.ID}
			return c.Post(p.s.Ctx(), "/groups/"+url.PathEscape(gid)+"/members/$ref", ref, nil)
		})
	}

	for _, tid := range req.TeamIDs {
		r.do("Add to team", tid, func() error {
			return c.Post(p.s.Ctx(), "/teams/"+url.PathEscape(tid)+"/members", conversationMember(req.Upn, false), nil)
		})
	}

	for _, cr := range req.ChannelRefs {
		cr := cr
		r.do("Add to channel", cr.TeamID+"/"+cr.ChannelID, func() error {
			return c.Post(p.s.Ctx(),
				"/teams/"+url.PathEscape(cr.TeamID)+"/channels/"+url.PathEscape(cr.ChannelID)+"/members",
				conversationMember(req.Upn, false), nil)
		})
	}

	p.s.Record("playbook.onboard", req.Upn, "steps="+itoa(len(r.steps)), nil)
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
	r := &runner{ok: true}
	u := url.PathEscape(req.Upn)

	if req.Block {
		r.do("Block sign-in", req.Upn, func() error {
			return c.Patch(p.s.Ctx(), "/users/"+u, map[string]any{"accountEnabled": false}, nil)
		})
	}
	if req.RevokeSessions {
		r.do("Revoke sessions", req.Upn, func() error {
			return c.Post(p.s.Ctx(), "/users/"+u+"/revokeSignInSessions", map[string]any{}, nil)
		})
	}
	if req.Oof {
		r.do("Set auto-reply (OOF)", req.Upn, func() error {
			msg := req.OofMessage
			if msg == "" {
				msg = "This employee is no longer with the organization. Your message will not be forwarded automatically."
			}
			return c.Patch(p.s.Ctx(), "/users/"+u+"/mailboxSettings", map[string]any{
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
			return c.Post(p.s.Ctx(), "/users/"+u+"/mailFolders/inbox/messageRules", map[string]any{
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
			return c.Patch(p.s.Ctx(), "/users/"+u, map[string]any{"showInAddressList": false}, nil)
		})
	}
	if req.CalendarTo != "" {
		r.do("Share calendar (read)", req.Upn+" → "+req.CalendarTo, func() error {
			return c.Post(p.s.Ctx(), "/users/"+u+"/calendar/calendarPermissions", map[string]any{
				"emailAddress": map[string]any{"address": req.CalendarTo},
				"role":         "read",
			}, nil)
		})
	}
	if req.BackupToUser != "" {
		drive := NewDriveService(p.s)
		r.do("Backup OneDrive", req.Upn+" → "+req.BackupToUser, func() error {
			folder := req.BackupFolder
			if folder == "" {
				folder = req.Upn
			}
			_, e := drive.CopyBetweenUsers(req.Upn, req.BackupToUser, folder, false)
			return e
		})
	}
	if req.RemoveFromGroups {
		// One report step per group so the operator sees exactly what happened.
		// Dynamic-membership and Exchange-managed (distribution) groups fail
		// individually with the Graph error; the rest still get removed.
		var idResp struct {
			ID string `json:"id"`
		}
		if err := c.Get(p.s.Ctx(), "/users/"+u, url.Values{"$select": {"id"}}, &idResp); err != nil {
			r.do("Remove from groups", req.Upn, func() error { return err })
		} else if items, err := c.ListAll(p.s.Ctx(), "/users/"+u+"/memberOf", url.Values{"$select": {"id,displayName"}}, 0); err != nil {
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
					return c.Delete(p.s.Ctx(), "/groups/"+url.PathEscape(gid)+"/members/"+url.PathEscape(idResp.ID)+"/$ref")
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
			if e := c.Get(p.s.Ctx(), "/users/"+u+"/licenseDetails", nil, &lic); e != nil {
				return e
			}
			remove := make([]string, 0, len(lic.Value))
			for _, l := range lic.Value {
				remove = append(remove, l.SkuID)
			}
			if len(remove) == 0 {
				return nil
			}
			return c.Post(p.s.Ctx(), "/users/"+u+"/assignLicense",
				map[string]any{"addLicenses": []any{}, "removeLicenses": remove}, nil)
		})
	}
	if req.Delete {
		r.do("Delete user", req.Upn, func() error {
			return c.Delete(p.s.Ctx(), "/users/"+u)
		})
	}

	p.s.Record("playbook.offboard", req.Upn, "steps="+itoa(len(r.steps)), nil)
	return r.result(), nil
}

var _ = json.Marshal
