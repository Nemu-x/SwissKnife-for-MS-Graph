package services

import (
	"encoding/json"
	"errors"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"

	"swissknife-app/internal/session"
)

// Exchange Online message trace over Microsoft Graph v1.0 (GA January 2026):
// GET /admin/exchange/tracing/messageTraces and
// GET /admin/exchange/tracing/messageTraces/{id}/getDetailsByRecipient(recipientAddress='...').
// Permission: ExchangeMessageTrace.Read.All (delegated or application).
// Source: https://learn.microsoft.com/en-us/graph/api/messagetracingroot-list-messagetraces?view=graph-rest-1.0
//
// Prerequisite: the tenant must hold a service principal for Microsoft's
// "Transport Data Platform" application (appId below). Without it Graph answers
// 401/403 even when the permission is consented. Creating it is a one-time write
// that needs Application.ReadWrite.All; the SP can take a few hours to become
// effective on Microsoft's side.
const messageTraceAppID = "8bd644d1-64a1-4d4b-ae52-2e0cbf64e373"

const messageTracesPath = "/admin/exchange/tracing/messageTraces"

// API limits documented for List messageTraces: the receivedDateTime interval
// must not exceed 10 days, the start must be within the last 90 days, and $top
// accepts 1..5000 (default page 1000). With no filter Graph returns the last 48h.
const (
	traceMaxDays     = 10
	traceDefaultDays = 2
	traceMaxTop      = 5000
	traceDefaultTop  = 100
)

type MailTraceService struct {
	s *session.Session
}

func NewMailTraceService(s *session.Session) *MailTraceService { return &MailTraceService{s: s} }

// TraceQuery is "the email did not arrive, where did it go?" — either side of
// the conversation can be empty; both may be external SMTP addresses.
type TraceQuery struct {
	Sender    string `json:"sender"`    // exact SMTP address of the purported sender
	Recipient string `json:"recipient"` // exact SMTP address the message was addressed to
	Days      int    `json:"days"`      // look back this many days; 0 = 2, capped at 10 by the API
	Top       int    `json:"top"`       // max rows; 0 = 100, capped at 5000
}

// Trace lists message traces matching the query. Both time bounds are always
// sent, as the docs require for a custom window.
func (m *MailTraceService) Trace(q TraceQuery) ([]json.RawMessage, error) {
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	days := q.Days
	if days <= 0 {
		days = traceDefaultDays
	}
	if days > traceMaxDays {
		days = traceMaxDays
	}
	top := q.Top
	if top <= 0 {
		top = traceDefaultTop
	}
	if top > traceMaxTop {
		top = traceMaxTop
	}
	sender := strings.TrimSpace(q.Sender)
	recipient := strings.TrimSpace(q.Recipient)
	if sender == "" && recipient == "" {
		return nil, errors.New("message trace: give a sender or a recipient address")
	}

	now := time.Now().UTC()
	const stamp = "2006-01-02T15:04:05Z"
	filters := []string{
		"receivedDateTime ge " + now.AddDate(0, 0, -days).Format(stamp),
		"receivedDateTime le " + now.Format(stamp),
	}
	if sender != "" {
		filters = append(filters, "senderAddress eq '"+escapeODataLiteral(sender)+"'")
	}
	if recipient != "" {
		filters = append(filters, "recipientAddress eq '"+escapeODataLiteral(recipient)+"'")
	}
	params := topParams(top)
	params.Set("$filter", strings.Join(filters, " and "))

	out, err := c.ListAll(m.s.Ctx(), messageTracesPath, params, top)
	target := sender
	if target == "" {
		target = recipient
	}
	m.s.Record("mail.trace", target, "sender="+sender+" recipient="+recipient+" days="+strconv.Itoa(days), err)
	return out, err
}

// Details returns the per-hop events (receive, deliver, defer, fail, quarantine…)
// of one traced message for one recipient. messageTraceID is the `id` of a row
// returned by Trace, not the Message-ID header.
func (m *MailTraceService) Details(messageTraceID, recipient string) ([]json.RawMessage, error) {
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	id := strings.TrimSpace(messageTraceID)
	recipient = strings.TrimSpace(recipient)
	if id == "" || recipient == "" {
		return nil, errors.New("message trace: both the trace id and the recipient address are required")
	}
	path := messageTracesPath + "/" + url.PathEscape(id) +
		"/getDetailsByRecipient(recipientAddress='" + url.PathEscape(escapeODataLiteral(recipient)) + "')"
	// The function returns a collection envelope; no paging is documented, but
	// following nextLink if present costs nothing.
	out, err := c.ListAll(m.s.Ctx(), path, nil, 0)
	m.s.Record("mail.trace.details", recipient, "traceId="+id, err)
	return out, err
}

// Prerequisite reports whether the Microsoft service principal the trace API
// depends on exists in this tenant. Needs Application.Read.All (or
// Directory.Read.All) to look.
func (m *MailTraceService) Prerequisite() (bool, error) {
	c, err := m.s.Client()
	if err != nil {
		return false, err
	}
	params := url.Values{
		"$filter": {"appId eq '" + messageTraceAppID + "'"},
		"$select": {"id"},
	}
	rows, err := c.ListAll(m.s.Ctx(), "/servicePrincipals", params, 1)
	if err != nil {
		return false, err
	}
	return len(rows) > 0, nil
}

// Provision creates the Microsoft service principal the trace API depends on.
// One-time tenant write; needs Application.ReadWrite.All.
func (m *MailTraceService) Provision() error {
	if err := m.s.GuardWrite(); err != nil {
		return err
	}
	c, err := m.s.Client()
	if err != nil {
		return err
	}
	body := map[string]string{"appId": messageTraceAppID}
	err = c.Do(m.s.Ctx(), http.MethodPost, "/servicePrincipals", nil, body, nil)
	m.s.Record("mail.trace.provision", messageTraceAppID, "create service principal for message trace", err)
	return err
}
