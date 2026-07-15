package services

import (
	"encoding/json"
	"net/url"
	"sort"
	"time"

	"swissknife-app/internal/session"
)

// AppsService — Entra app registrations, with credential-expiry monitoring.
type AppsService struct {
	s *session.Session
}

func NewAppsService(s *session.Session) *AppsService { return &AppsService{s: s} }

func (a *AppsService) List(search string, maxItems int) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{
		"$top":    {"100"},
		"$select": {"id,appId,displayName,signInAudience,createdDateTime,passwordCredentials,keyCredentials"},
	}
	if search != "" {
		params.Set("$filter", "startswith(displayName,'"+escapeODataLiteral(search)+"')")
	}
	return c.ListAll(a.s.Ctx(), "/applications", params, maxItems)
}

// AddSecret issues a new client secret on an app registration (rotation).
// objectID is the application OBJECT id, not the appId. The secret text is
// returned once and never stored anywhere.
func (a *AppsService) AddSecret(objectID, displayName string, months int) (map[string]any, error) {
	if err := a.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if months < 1 || months > 24 {
		months = 6
	}
	if displayName == "" {
		displayName = "SwissKnife rotation " + time.Now().Format("2006-01-02")
	}
	var resp map[string]any
	err = c.Post(a.s.Ctx(), "/applications/"+url.PathEscape(objectID)+"/addPassword", map[string]any{
		"passwordCredential": map[string]any{
			"displayName": displayName,
			"endDateTime": time.Now().AddDate(0, months, 0).UTC().Format(time.RFC3339),
		},
	}, &resp)
	a.s.Record("apps.addSecret", objectID, displayName, err)
	if err != nil {
		return nil, err
	}
	return resp, nil
}

// ExpiringCredential is one secret/cert nearing or past expiry.
type ExpiringCredential struct {
	AppName     string `json:"appName"`
	AppID       string `json:"appId"`
	Kind        string `json:"kind"` // secret | certificate
	DisplayName string `json:"displayName"`
	Expires     string `json:"expires"`
	DaysLeft    int    `json:"daysLeft"`
}

// Expiring returns app secrets/certificates expiring within `days` (or already
// expired, with negative daysLeft), sorted by soonest first.
func (a *AppsService) Expiring(days int) ([]ExpiringCredential, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	if days <= 0 {
		days = 30
	}
	apps, err := c.ListAll(a.s.Ctx(), "/applications",
		url.Values{"$top": {"100"}, "$select": {"appId,displayName,passwordCredentials,keyCredentials"}}, 0)
	if err != nil {
		return nil, err
	}

	cutoff := time.Now().AddDate(0, 0, days)
	var out []ExpiringCredential
	for _, raw := range apps {
		var app struct {
			AppID       string `json:"appId"`
			DisplayName string `json:"displayName"`
			Passwords   []struct {
				DisplayName string    `json:"displayName"`
				EndDateTime time.Time `json:"endDateTime"`
			} `json:"passwordCredentials"`
			Keys []struct {
				DisplayName string    `json:"displayName"`
				EndDateTime time.Time `json:"endDateTime"`
			} `json:"keyCredentials"`
		}
		if json.Unmarshal(raw, &app) != nil {
			continue
		}
		add := func(kind, name string, exp time.Time) {
			if exp.IsZero() || exp.After(cutoff) {
				return
			}
			out = append(out, ExpiringCredential{
				AppName:     app.DisplayName,
				AppID:       app.AppID,
				Kind:        kind,
				DisplayName: name,
				Expires:     exp.Format("2006-01-02"),
				DaysLeft:    int(time.Until(exp).Hours() / 24),
			})
		}
		for _, p := range app.Passwords {
			add("secret", p.DisplayName, p.EndDateTime)
		}
		for _, k := range app.Keys {
			add("certificate", k.DisplayName, k.EndDateTime)
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i].DaysLeft < out[j].DaysLeft })
	return out, nil
}
