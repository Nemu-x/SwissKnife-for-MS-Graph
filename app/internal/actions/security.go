package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"sort"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// Incident response on Graph: the inbox-rule audit (attackers hide their
// tracks with rules that forward mail out or delete replies) and reporting a
// message to Microsoft as phishing.

func securityActions() []engine.Action {
	return []engine.Action{
		{Manifest: engine.Manifest{ID: "mail.ruleAudit", Page: "security", Danger: engine.Read,
			Fields:      []engine.Field{{Name: "user", Kind: engine.FieldUser}},
			Permissions: []string{"MailboxSettings.Read", "User.Read.All", "Domain.Read.All"}},
			Impls: []engine.Impl{engine.ReadImpl(graphRuleAudit{})}},
		{Manifest: engine.Manifest{ID: "mail.reportThreat", Page: "audit", Danger: engine.Write,
			Fields: []engine.Field{
				{Name: "user", Kind: engine.FieldUser, Required: true},
				{Name: "sender", Kind: engine.FieldText, Required: true},
				{Name: "category", Kind: engine.FieldChoice, Required: true, Options: []string{"phishing", "spam", "malware", "notJunk"}, Default: "phishing"},
			},
			Permissions: []string{"ThreatSubmission.ReadWrite.All", "Mail.ReadBasic.All"}},
			Impls: []engine.Impl{graphReportThreat{}}},
	}
}

// ruleScanConcurrency bounds parallel mailbox reads (Graph throttles per app).
const ruleScanConcurrency = 8

type graphRuleAudit struct{}

func (graphRuleAudit) Backend() engine.Backend { return engine.BackendGraph }

type messageRule struct {
	DisplayName string `json:"displayName"`
	IsEnabled   bool   `json:"isEnabled"`
	Actions     struct {
		Delete              bool   `json:"delete"`
		PermanentDelete     bool   `json:"permanentDelete"`
		MarkAsRead          bool   `json:"markAsRead"`
		MoveToFolder        string `json:"moveToFolder"`
		ForwardTo           []recp `json:"forwardTo"`
		ForwardAsAttachment []recp `json:"forwardAsAttachmentTo"`
		RedirectTo          []recp `json:"redirectTo"`
	} `json:"actions"`
}

type recp struct {
	EmailAddress struct {
		Address string `json:"address"`
	} `json:"emailAddress"`
}

// suspicious explains why a rule looks like an attacker's, or "" when it
// does not: mail leaving the tenant, or mail made to disappear. The reasons
// are tokens ("externalForward=addr; deletes") the UI translates.
func suspicious(r messageRule, internal map[string]bool, hidden map[string]bool) string {
	var why []string
	for _, list := range [][]recp{r.Actions.ForwardTo, r.Actions.ForwardAsAttachment, r.Actions.RedirectTo} {
		for _, x := range list {
			addr := strings.ToLower(x.EmailAddress.Address)
			if at := strings.LastIndex(addr, "@"); at >= 0 && !internal[addr[at+1:]] {
				why = append(why, "externalForward="+addr)
			}
		}
	}
	if r.Actions.Delete || r.Actions.PermanentDelete {
		why = append(why, "deletes")
	}
	if r.Actions.MoveToFolder != "" && hidden[r.Actions.MoveToFolder] && r.Actions.MarkAsRead {
		why = append(why, "hides")
	}
	return strings.Join(why, "; ")
}

func (graphRuleAudit) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	var domains struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	if err := env.Graph.Get(env.Ctx, "/domains", url.Values{"$select": {"id"}}, &domains); err != nil {
		return nil, err
	}
	internal := map[string]bool{}
	for _, d := range domains.Value {
		internal[strings.ToLower(d.ID)] = true
	}

	var users []string
	if in["user"] != "" {
		u, err := getUser(env, in["user"], "id,userPrincipalName")
		if err != nil {
			return nil, err
		}
		users = []string{u.UPN}
	} else {
		all, err := graphapi.ListAllInto[struct {
			UPN  string `json:"userPrincipalName"`
			Mail string `json:"mail"`
		}](env.Ctx, env.Graph, "/users", url.Values{"$select": {"userPrincipalName,mail"}, "$top": {"999"}}, 0)
		if err != nil {
			return nil, err
		}
		for _, u := range all {
			if u.Mail != "" { // no mail, no mailbox to scan
				users = append(users, u.UPN)
			}
		}
	}

	type found struct {
		user string
		rule messageRule
		why  string
	}
	var (
		mu      sync.Mutex
		hits    []found
		skipped int
		wg      sync.WaitGroup
		sem     = make(chan struct{}, ruleScanConcurrency)
	)
	for _, upn := range users {
		if env.Ctx.Err() != nil {
			break
		}
		wg.Add(1)
		sem <- struct{}{}
		go func(upn string) {
			defer func() { <-sem; wg.Done() }()
			base := "/users/" + url.PathEscape(upn) + "/mailFolders"
			// Well-known folders the "hide it" pattern moves mail into, by id.
			hidden := map[string]bool{}
			for _, wk := range []string{"deleteditems", "junkemail", "archive", "conversationhistory"} {
				var f struct {
					ID string `json:"id"`
				}
				if env.Graph.Get(env.Ctx, base+"/"+wk, url.Values{"$select": {"id"}}, &f) == nil && f.ID != "" {
					hidden[f.ID] = true
				}
			}
			rules, err := graphapi.ListAllInto[messageRule](env.Ctx, env.Graph, base+"/inbox/messageRules", nil, 0)
			mu.Lock()
			defer mu.Unlock()
			if err != nil {
				skipped++ // no mailbox, or no access to it
				return
			}
			for _, r := range rules {
				if why := suspicious(r, internal, hidden); why != "" {
					hits = append(hits, found{upn, r, why})
				}
			}
		}(upn)
	}
	wg.Wait()
	if err := env.Ctx.Err(); err != nil {
		return nil, err
	}

	sort.Slice(hits, func(i, j int) bool { return hits[i].user < hits[j].user })
	res := &engine.ReadResult{Columns: []string{"user", "rule", "enabled", "why"}}
	for _, h := range hits {
		res.Rows = append(res.Rows, engine.Row{"user": h.user, "rule": h.rule.DisplayName, "enabled": yesNo(h.rule.IsEnabled), "why": h.why})
	}
	if len(users) > 1 {
		res.Note = &engine.Reason{Key: "rulesScanned", Params: map[string]string{
			"scanned": fmt.Sprint(len(users) - skipped), "unreadable": fmt.Sprint(skipped)}}
	}
	return res, nil
}

func yesNo(b bool) string {
	if b {
		return "yes"
	}
	return "no"
}

// --- mail.reportThreat ------------------------------------------------------------

type graphReportThreat struct{}

func (graphReportThreat) Backend() engine.Backend { return engine.BackendGraph }

func (graphReportThreat) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	sender := strings.ToLower(strings.TrimSpace(in["sender"]))
	if !emailRe.MatchString(sender) {
		return nil, errors.New("sender must be the full address the message came from")
	}
	u, err := getUser(env, in["user"], "id,userPrincipalName")
	if err != nil {
		return nil, err
	}
	// The newest message from the sender in the recipient's mailbox.
	var msgs struct {
		Value []struct {
			ID       string    `json:"id"`
			Subject  string    `json:"subject"`
			Received time.Time `json:"receivedDateTime"`
		} `json:"value"`
	}
	q := url.Values{
		"$filter":  {"from/emailAddress/address eq '" + escapeOData(sender) + "'"},
		"$select":  {"id,subject,receivedDateTime"},
		"$orderby": {"receivedDateTime desc"},
		"$top":     {"1"},
	}
	if err := env.Graph.Get(env.Ctx, "/users/"+url.PathEscape(u.ID)+"/messages", q, &msgs); err != nil {
		return nil, err
	}
	if len(msgs.Value) == 0 {
		return nil, fmt.Errorf("no message from %s in %s's mailbox", sender, u.UPN)
	}
	m := msgs.Value[0]
	return []engine.Change{{
		Target: fmt.Sprintf("%q (%s)", m.Subject, m.Received.Format("2006-01-02 15:04")),
		Field:  "threatReport", Op: "add", After: in["category"],
		Ref: map[string]string{"url": "https://graph.microsoft.com/v1.0/users/" + u.ID + "/messages/" + m.ID, "recipient": u.UPN},
	}}, nil
}

func (graphReportThreat) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	body := map[string]any{
		"@odata.type":           "#microsoft.graph.security.emailUrlThreatSubmission",
		"category":              in["category"],
		"recipientEmailAddress": ch.Ref["recipient"],
		"messageUrl":            ch.Ref["url"],
	}
	var out json.RawMessage
	return env.Graph.Post(env.Ctx, "/security/threatSubmission/emailThreats", body, &out)
}

func escapeOData(s string) string { return strings.ReplaceAll(s, "'", "''") }
