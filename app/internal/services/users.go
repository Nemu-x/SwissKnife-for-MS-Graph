// Package services holds domain services bound into Wails; the frontend calls only these.
package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

type UsersService struct {
	s *session.Session
}

func NewUsersService(s *session.Session) *UsersService { return &UsersService{s: s} }

func topParams(top int) url.Values {
	if top <= 0 {
		top = 50
	}
	return url.Values{"$top": {strconv.Itoa(top)}}
}

// List returns tenant users; maxItems=0 means all pages.
func (u *UsersService) List(search string, maxItems int) ([]json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{
		"$top":    {"100"},
		"$select": {"id,displayName,userPrincipalName,mail,accountEnabled,jobTitle,department"},
	}
	if search != "" {
		params.Set("$filter", "startswith(displayName,'"+escapeODataLiteral(search)+"') or startswith(userPrincipalName,'"+escapeODataLiteral(search)+"')")
	}
	return c.ListAll(u.s.Ctx(), "/users", params, maxItems)
}

func (u *UsersService) Get(user string) (json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(u.s.Ctx(), "/users/"+url.PathEscape(user), nil, &out)
	return out, err
}

func (u *UsersService) MemberOf(user string) ([]json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(u.s.Ctx(), "/users/"+url.PathEscape(user)+"/memberOf", nil, 0)
}

func (u *UsersService) LicenseDetails(user string) ([]json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(u.s.Ctx(), "/users/"+url.PathEscape(user)+"/licenseDetails", nil, 0)
}

// Snapshot is a per-user summary: profile + groups + licenses.
func (u *UsersService) Snapshot(user string) (map[string]any, error) {
	profile, err := u.Get(user)
	if err != nil {
		return nil, err
	}
	groups, err := u.MemberOf(user)
	if err != nil {
		return nil, err
	}
	licenses, err := u.LicenseDetails(user)
	if err != nil {
		return nil, err
	}
	return map[string]any{
		"profile":  profile,
		"memberOf": groups,
		"licenses": licenses,
	}, nil
}

func (u *UsersService) setEnabled(user string, enabled bool) error {
	action := "users.block"
	if enabled {
		action = "users.unblock"
	}
	if err := u.s.GuardWrite(); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	err = c.Patch(u.s.Ctx(), "/users/"+url.PathEscape(user), map[string]any{"accountEnabled": enabled}, nil)
	u.s.Record(action, user, "", err)
	return err
}

func (u *UsersService) Block(user string) error   { return u.setEnabled(user, false) }
func (u *UsersService) Unblock(user string) error { return u.setEnabled(user, true) }

// ResetPassword is destructive: requires typed confirm (entering the target UPN).
func (u *UsersService) ResetPassword(user, newPassword string, forceChange bool, confirm string) error {
	if err := u.s.GuardDestructive(user, confirm); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	body := map[string]any{
		"passwordProfile": map[string]any{
			"forceChangePasswordNextSignIn": forceChange,
			"password":                      newPassword,
		},
	}
	err = c.Patch(u.s.Ctx(), "/users/"+url.PathEscape(user), body, nil)
	u.s.Record("users.resetPassword", user, "forceChange="+strconv.FormatBool(forceChange), err)
	return err
}

func (u *UsersService) RevokeSessions(user, confirm string) error {
	if err := u.s.GuardDestructive(user, confirm); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	err = c.Post(u.s.Ctx(), "/users/"+url.PathEscape(user)+"/revokeSignInSessions", map[string]any{}, nil)
	u.s.Record("users.revokeSessions", user, "", err)
	return err
}

// escapeODataLiteral doubles single quotes inside an OData literal.
func escapeODataLiteral(s string) string {
	out := make([]rune, 0, len(s))
	for _, r := range s {
		if r == '\'' {
			out = append(out, '\'', '\'')
		} else {
			out = append(out, r)
		}
	}
	return string(out)
}
