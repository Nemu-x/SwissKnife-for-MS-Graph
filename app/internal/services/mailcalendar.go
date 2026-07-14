package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

type MailService struct {
	s *session.Session
}

func NewMailService(s *session.Session) *MailService { return &MailService{s: s} }

func (m *MailService) List(user, folder string, top int) ([]json.RawMessage, error) {
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	if folder == "" {
		folder = "inbox"
	}
	if top <= 0 {
		top = 25
	}
	params := url.Values{
		"$top":     {strconv.Itoa(top)},
		"$select":  {"id,subject,from,receivedDateTime,isRead,hasAttachments"},
		"$orderby": {"receivedDateTime DESC"},
	}
	return c.ListAll(m.s.Ctx(),
		"/users/"+url.PathEscape(user)+"/mailFolders/"+url.PathEscape(folder)+"/messages",
		params, top)
}

// Send is destructive (sending as any user): typed confirm on the UPN.
func (m *MailService) Send(user, subject, bodyText string, to []string, confirm string) error {
	if err := m.s.GuardDestructive(user, confirm); err != nil {
		return err
	}
	c, err := m.s.Client()
	if err != nil {
		return err
	}
	recipients := make([]map[string]any, 0, len(to))
	for _, addr := range to {
		recipients = append(recipients, map[string]any{"emailAddress": map[string]any{"address": addr}})
	}
	payload := map[string]any{
		"message": map[string]any{
			"subject":      subject,
			"body":         map[string]any{"contentType": "Text", "content": bodyText},
			"toRecipients": recipients,
		},
		"saveToSentItems": true,
	}
	err = c.Post(m.s.Ctx(), "/users/"+url.PathEscape(user)+"/sendMail", payload, nil)
	m.s.Record("mail.send", user, "to="+strconv.Itoa(len(to))+" recipients", err)
	return err
}

type CalendarService struct {
	s *session.Session
}

func NewCalendarService(s *session.Session) *CalendarService { return &CalendarService{s: s} }

func (cal *CalendarService) List(user string, top int) ([]json.RawMessage, error) {
	c, err := cal.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 25
	}
	params := url.Values{
		"$top":     {strconv.Itoa(top)},
		"$orderby": {"start/dateTime DESC"},
		"$select":  {"id,subject,start,end,location,organizer,attendees"},
	}
	return c.ListAll(cal.s.Ctx(), "/users/"+url.PathEscape(user)+"/events", params, top)
}

func (cal *CalendarService) CreateEvent(user, subject, bodyText, startISO, endISO, timezone string, attendees []string) (json.RawMessage, error) {
	if err := cal.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := cal.s.Client()
	if err != nil {
		return nil, err
	}
	att := make([]map[string]any, 0, len(attendees))
	for _, a := range attendees {
		att = append(att, map[string]any{
			"emailAddress": map[string]any{"address": a},
			"type":         "required",
		})
	}
	event := map[string]any{
		"subject":   subject,
		"body":      map[string]any{"contentType": "Text", "content": bodyText},
		"start":     map[string]any{"dateTime": startISO, "timeZone": timezone},
		"end":       map[string]any{"dateTime": endISO, "timeZone": timezone},
		"attendees": att,
	}
	var out json.RawMessage
	err = c.Post(cal.s.Ctx(), "/users/"+url.PathEscape(user)+"/events", event, &out)
	cal.s.Record("calendar.createEvent", user, "subject="+subject, err)
	return out, err
}
