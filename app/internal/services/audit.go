package services

import (
	"encoding/json"
	"strconv"
	"strings"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/session"
)

type AuditService struct {
	s *session.Session
}

func NewAuditService(s *session.Session) *AuditService { return &AuditService{s: s} }

func (a *AuditService) DirectoryAudits(top int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	return c.ListAll(a.s.Ctx(), "/auditLogs/directoryAudits", topParams(top), top)
}

func (a *AuditService) SignIns(top int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	return c.ListAll(a.s.Ctx(), "/auditLogs/signIns", topParams(top), top)
}

// SignInQuery narrows the sign-in log to the question actually being asked:
// "why can this person not sign in", not "show me the last 50 events".
type SignInQuery struct {
	Upn        string `json:"upn"`        // exact UPN, empty for the whole tenant
	Days       int    `json:"days"`       // look back this many days, 0 = no limit
	FailedOnly bool   `json:"failedOnly"` // only sign-ins Entra rejected
	Top        int    `json:"top"`
}

// SignInsFiltered is the filtered sign-in log. Graph supports $filter on
// userPrincipalName, createdDateTime and status/errorCode for this collection.
func (a *AuditService) SignInsFiltered(q SignInQuery) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	top := q.Top
	if top <= 0 {
		top = 50
	}
	params := topParams(top)
	filters := []string{}
	if upn := strings.TrimSpace(q.Upn); upn != "" {
		filters = append(filters, "userPrincipalName eq '"+escapeODataLiteral(upn)+"'")
	}
	if q.Days > 0 {
		since := time.Now().UTC().AddDate(0, 0, -q.Days).Format("2006-01-02T15:04:05Z")
		filters = append(filters, "createdDateTime ge "+since)
	}
	if q.FailedOnly {
		// 0 is "success"; anything else is a failure or an interrupt.
		filters = append(filters, "status/errorCode ne 0")
	}
	if len(filters) > 0 {
		params.Set("$filter", strings.Join(filters, " and "))
	}
	out, err := c.ListAll(a.s.Ctx(), "/auditLogs/signIns", params, top)
	a.s.Record("audit.signIns", q.Upn, "days="+strconv.Itoa(q.Days)+" failedOnly="+strconv.FormatBool(q.FailedOnly), err)
	return out, err
}

// DirectoryAuditsFiltered narrows directory changes to one target or actor and a
// time window — the "who touched this account" question.
func (a *AuditService) DirectoryAuditsFiltered(search string, days, top int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	params := topParams(top)
	filters := []string{}
	if days > 0 {
		since := time.Now().UTC().AddDate(0, 0, -days).Format("2006-01-02T15:04:05Z")
		filters = append(filters, "activityDateTime ge "+since)
	}
	if s := strings.TrimSpace(search); s != "" {
		// Graph cannot filter directoryAudits by target, so the actor is the one
		// filterable side; the UI says so.
		filters = append(filters, "initiatedBy/user/userPrincipalName eq '"+escapeODataLiteral(s)+"'")
	}
	if len(filters) > 0 {
		params.Set("$filter", strings.Join(filters, " and "))
	}
	out, err := c.ListAll(a.s.Ctx(), "/auditLogs/directoryAudits", params, top)
	a.s.Record("audit.directory", search, "days="+strconv.Itoa(days), err)
	return out, err
}

// Activity is the local app action log (ADR-002).
func (a *AuditService) Activity(n int) []auditlog.Entry {
	if n <= 0 {
		n = 200
	}
	return a.s.Audit.Tail(n)
}
