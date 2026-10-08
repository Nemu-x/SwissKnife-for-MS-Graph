package actions

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"net/url"
	"sort"
	"strconv"
	"strings"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// The report catalog: the questions admins ask most ("who has not signed in",
// "which licenses are wasted", "who has no MFA"…) as read actions. Each one
// answers with a table the UI can export to CSV.

var reportDays = engine.Field{Name: "days", Kind: engine.FieldChoice, Required: true, Options: []string{"30", "90", "180"}, Default: "90"}

func reportActions() []engine.Action {
	read := func(id string, fields []engine.Field, perms []string, r engine.Reader) engine.Action {
		return engine.Action{Manifest: engine.Manifest{ID: id, Page: "reports", Danger: engine.Read, Fields: fields, Permissions: perms},
			Impls: []engine.Impl{engine.ReadImpl(r)}}
	}
	return []engine.Action{
		read("report.inactiveUsers", []engine.Field{reportDays}, []string{"User.Read.All", "AuditLog.Read.All"}, graphInactive{}),
		read("report.licenseWaste", []engine.Field{reportDays}, []string{"User.Read.All", "AuditLog.Read.All", "Organization.Read.All"}, graphInactive{licensedOnly: true}),
		read("report.guests", nil, []string{"User.Read.All", "AuditLog.Read.All"}, graphGuests{}),
		read("report.mfaStatus", nil, []string{"AuditLog.Read.All", "UserAuthenticationMethod.Read.All"}, graphMfaStatus{}),
		read("report.privilegedRoles", nil, []string{"RoleManagement.Read.Directory"}, graphPrivileged{}),
		read("report.mailboxSizes", nil, []string{"Reports.Read.All"}, graphMailboxSizes{}),
		{Manifest: engine.Manifest{ID: "report.mailboxForwarding", Page: "reports", Danger: engine.Read,
			Permissions: []string{"Exchange.ManageAsApp + View-Only Recipients"}},
			Impls: []engine.Impl{engine.ReadImpl(psForwardingReport{})}},
		{Manifest: engine.Manifest{ID: "mailbox.statistics", Page: "mail", Danger: engine.Read,
			Fields:      []engine.Field{{Name: "mailbox", Kind: engine.FieldUser, Required: true}},
			Permissions: []string{"Exchange.ManageAsApp + View-Only Recipients"}},
			Impls: []engine.Impl{engine.ReadImpl(psMailboxStatistics{})}},
	}
}

// reportCmdlets join the Exchange allow-list.
var reportCmdlets = []string{"Get-MailboxStatistics"}

func dateOnly(t *time.Time) string {
	if t == nil || t.IsZero() {
		return ""
	}
	return t.Format("2006-01-02")
}

// --- inactive users / license waste ---------------------------------------------------

type graphInactive struct {
	licensedOnly bool
}

func (graphInactive) Backend() engine.Backend { return engine.BackendGraph }

func (g graphInactive) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	days, _ := strconv.Atoi(in["days"])
	cutoff := time.Now().AddDate(0, 0, -days).UTC().Format("2006-01-02T15:04:05Z")
	type user struct {
		UPN            string `json:"userPrincipalName"`
		AccountEnabled bool   `json:"accountEnabled"`
		Licenses       []struct {
			SkuID string `json:"skuId"`
		} `json:"assignedLicenses"`
		SignIn struct {
			Last *time.Time `json:"lastSignInDateTime"`
		} `json:"signInActivity"`
	}
	// Users who never signed in have no lastSignInDateTime and do not match
	// "le": the note says so instead of mixing them in.
	users, err := graphapi.ListAllInto[user](env.Ctx, env.Graph, "/users", url.Values{
		"$select": {"userPrincipalName,accountEnabled,assignedLicenses,signInActivity"},
		"$filter": {"signInActivity/lastSignInDateTime le " + cutoff},
	}, 0)
	if err != nil {
		return nil, err
	}
	skus := map[string]string{}
	if g.licensedOnly {
		raw, err := env.Graph.ListAll(env.Ctx, "/subscribedSkus", url.Values{"$select": {"skuId,skuPartNumber"}}, 0)
		if err != nil {
			return nil, err
		}
		for _, r := range raw {
			var s struct{ SkuID, SkuPartNumber string }
			if json.Unmarshal(r, &s) == nil {
				skus[s.SkuID] = s.SkuPartNumber
			}
		}
	}
	res := &engine.ReadResult{Columns: []string{"user", "lastSignIn", "enabled"}}
	if g.licensedOnly {
		res.Columns = []string{"user", "lastSignIn", "licenses", "enabled"}
	}
	sort.Slice(users, func(i, j int) bool { return dateOnly(users[i].SignIn.Last) < dateOnly(users[j].SignIn.Last) })
	for _, u := range users {
		row := engine.Row{"user": u.UPN, "lastSignIn": dateOnly(u.SignIn.Last), "enabled": yesNo(u.AccountEnabled)}
		if g.licensedOnly {
			if len(u.Licenses) == 0 {
				continue
			}
			names := make([]string, 0, len(u.Licenses))
			for _, l := range u.Licenses {
				names = append(names, firstOf(skus[l.SkuID], l.SkuID))
			}
			row["licenses"] = strings.Join(names, ", ")
		}
		res.Rows = append(res.Rows, row)
	}
	res.Note = &engine.Reason{Key: "neverSignedInExcluded"}
	return res, nil
}

// --- guests --------------------------------------------------------------------------

type graphGuests struct{}

func (graphGuests) Backend() engine.Backend { return engine.BackendGraph }

func (graphGuests) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	type guest struct {
		Mail    string     `json:"mail"`
		Name    string     `json:"displayName"`
		Created *time.Time `json:"createdDateTime"`
		State   string     `json:"externalUserState"`
		SignIn  struct {
			Last *time.Time `json:"lastSignInDateTime"`
		} `json:"signInActivity"`
	}
	guests, err := graphapi.ListAllInto[guest](env.Ctx, env.Graph, "/users", url.Values{
		"$select": {"mail,displayName,createdDateTime,externalUserState,signInActivity"},
		"$filter": {"userType eq 'Guest'"},
	}, 0)
	if err != nil {
		return nil, err
	}
	sort.Slice(guests, func(i, j int) bool { return dateOnly(guests[i].SignIn.Last) < dateOnly(guests[j].SignIn.Last) })
	res := &engine.ReadResult{Columns: []string{"guest", "name", "invited", "state", "lastSignIn"}}
	for _, g := range guests {
		res.Rows = append(res.Rows, engine.Row{"guest": g.Mail, "name": g.Name, "invited": dateOnly(g.Created),
			"state": g.State, "lastSignIn": dateOnly(g.SignIn.Last)})
	}
	return res, nil
}

// --- MFA status ----------------------------------------------------------------------

type graphMfaStatus struct{}

func (graphMfaStatus) Backend() engine.Backend { return engine.BackendGraph }

func (graphMfaStatus) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	type detail struct {
		UPN     string   `json:"userPrincipalName"`
		IsAdmin bool     `json:"isAdmin"`
		Methods []string `json:"methodsRegistered"`
		Type    string   `json:"userType"`
	}
	rows, err := graphapi.ListAllInto[detail](env.Ctx, env.Graph, "/reports/authenticationMethods/userRegistrationDetails", url.Values{
		"$filter": {"isMfaRegistered eq false"},
	}, 0)
	if err != nil {
		return nil, err
	}
	// Admins without MFA first: they are the expensive ones.
	sort.SliceStable(rows, func(i, j int) bool { return rows[i].IsAdmin && !rows[j].IsAdmin })
	res := &engine.ReadResult{Columns: []string{"user", "admin", "methods"}}
	for _, r := range rows {
		if strings.EqualFold(r.Type, "guest") {
			continue
		}
		res.Rows = append(res.Rows, engine.Row{"user": r.UPN, "admin": yesNo(r.IsAdmin), "methods": strings.Join(r.Methods, ", ")})
	}
	return res, nil
}

// --- privileged roles -------------------------------------------------------------------

type graphPrivileged struct{}

func (graphPrivileged) Backend() engine.Backend { return engine.BackendGraph }

func (graphPrivileged) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	type role struct {
		ID   string `json:"id"`
		Name string `json:"displayName"`
	}
	type member struct {
		Type string `json:"@odata.type"`
		UPN  string `json:"userPrincipalName"`
		Name string `json:"displayName"`
	}
	roles, err := graphapi.ListAllInto[role](env.Ctx, env.Graph, "/directoryRoles", url.Values{"$select": {"id,displayName"}}, 0)
	if err != nil {
		return nil, err
	}
	sort.Slice(roles, func(i, j int) bool { return roles[i].Name < roles[j].Name })
	res := &engine.ReadResult{Columns: []string{"role", "member", "kind"}}
	for _, r := range roles {
		// Members per role, paged: an expanded collection can be cut short.
		members, err := graphapi.ListAllInto[member](env.Ctx, env.Graph, "/directoryRoles/"+url.PathEscape(r.ID)+"/members", nil, 0)
		if err != nil {
			return nil, err
		}
		for _, m := range members {
			kind := "user"
			if strings.Contains(m.Type, "servicePrincipal") {
				kind = "app"
			} else if strings.Contains(m.Type, "group") {
				kind = "group"
			}
			res.Rows = append(res.Rows, engine.Row{"role": r.Name, "member": firstOf(m.UPN, m.Name), "kind": kind})
		}
	}
	return res, nil
}

// --- mailbox sizes -----------------------------------------------------------------------

type graphMailboxSizes struct{}

func (graphMailboxSizes) Backend() engine.Backend { return engine.BackendGraph }

func (graphMailboxSizes) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	text, err := env.Graph.GetText(env.Ctx, "/reports/getMailboxUsageDetail(period='D7')", nil)
	if err != nil {
		return nil, err
	}
	records, err := csv.NewReader(strings.NewReader(strings.TrimPrefix(text, string(rune(0xFEFF))))).ReadAll()
	if err != nil {
		return nil, fmt.Errorf("mailbox usage report: %w", err)
	}
	if len(records) < 1 {
		return &engine.ReadResult{}, nil
	}
	col := map[string]int{}
	for i, h := range records[0] {
		col[h] = i
	}
	get := func(r []string, name string) string {
		if i, ok := col[name]; ok && i < len(r) {
			return r[i]
		}
		return ""
	}
	type mbx struct {
		upn, items, last string
		bytes, quota     int64
	}
	var all []mbx
	for _, r := range records[1:] {
		b, _ := strconv.ParseInt(get(r, "Storage Used (Byte)"), 10, 64)
		q, _ := strconv.ParseInt(get(r, "Prohibit Send/Receive Quota (Byte)"), 10, 64)
		all = append(all, mbx{upn: get(r, "User Principal Name"), items: get(r, "Item Count"), last: get(r, "Last Activity Date"), bytes: b, quota: q})
	}
	sort.Slice(all, func(i, j int) bool { return all[i].bytes > all[j].bytes })
	res := &engine.ReadResult{Columns: []string{"mailbox", "used", "full", "items", "lastActivity"}}
	for _, m := range all {
		full := ""
		if m.quota > 0 {
			full = fmt.Sprintf("%d%%", m.bytes*100/m.quota)
		}
		res.Rows = append(res.Rows, engine.Row{"mailbox": m.upn, "used": humanSize(m.bytes), "full": full, "items": m.items, "lastActivity": m.last})
	}
	return res, nil
}

func humanSize(b int64) string {
	const unit = 1024
	if b < unit {
		return fmt.Sprintf("%d B", b)
	}
	div, exp := int64(unit), 0
	for n := b / unit; n >= unit; n /= unit {
		div *= unit
		exp++
	}
	return fmt.Sprintf("%.1f %cB", float64(b)/float64(div), "KMGTPE"[exp])
}

// --- Exchange: forwarding mailboxes, statistics --------------------------------------------

type psForwardingReport struct{ psBase }

func (psForwardingReport) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	rows, err := exo(env, "Get-Mailbox", map[string]any{
		"ResultSize": "Unlimited",
		"Filter":     "ForwardingSmtpAddress -ne $null -or ForwardingAddress -ne $null",
	}, "UserPrincipalName", "ForwardingSmtpAddress", "ForwardingAddress", "DeliverToMailboxAndForward")
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"mailbox", "forwardsTo", "keepsCopy"}}
	for _, raw := range rows {
		var m struct {
			UserPrincipalName, ForwardingSmtpAddress, ForwardingAddress string
			DeliverToMailboxAndForward                                  bool
		}
		if json.Unmarshal(raw, &m) == nil {
			to := firstOf(strings.TrimPrefix(strings.TrimPrefix(m.ForwardingSmtpAddress, "smtp:"), "SMTP:"), m.ForwardingAddress)
			res.Rows = append(res.Rows, engine.Row{"mailbox": m.UserPrincipalName, "forwardsTo": to, "keepsCopy": yesNo(m.DeliverToMailboxAndForward)})
		}
	}
	return res, nil
}

type psMailboxStatistics struct{ psBase }

func (psMailboxStatistics) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	rows, err := exo(env, "Get-MailboxStatistics", map[string]any{"Identity": owner.UPN},
		"TotalItemSize", "ItemCount", "DeletedItemCount", "TotalDeletedItemSize", "LastUserActionTime")
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"metric", "value"}}
	if len(rows) == 0 {
		return res, nil
	}
	var s map[string]json.RawMessage
	_ = json.Unmarshal(rows[0], &s)
	for _, k := range []string{"TotalItemSize", "ItemCount", "DeletedItemCount", "TotalDeletedItemSize", "LastUserActionTime"} {
		v := strings.Trim(string(s[k]), `"`)
		if v == "null" {
			v = ""
		}
		res.Rows = append(res.Rows, engine.Row{"metric": k, "value": v})
	}
	return res, nil
}
