package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

// RolesService — Entra directory (admin) roles.
type RolesService struct {
	s *session.Session
}

func NewRolesService(s *session.Session) *RolesService { return &RolesService{s: s} }

// List returns activated directory roles in the tenant.
func (r *RolesService) List() ([]json.RawMessage, error) {
	c, err := r.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(r.s.Ctx(), "/directoryRoles", nil, 0)
}

// Members lists members of a directory role.
func (r *RolesService) Members(roleID string) ([]json.RawMessage, error) {
	c, err := r.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(r.s.Ctx(), "/directoryRoles/"+url.PathEscape(roleID)+"/members", nil, 0)
}

// AddMember assigns a user (by UPN/id) to a directory role. Destructive-adjacent
// (grants privilege) so it is write-guarded and audited.
func (r *RolesService) AddMember(roleID, upn string) error {
	if err := r.s.GuardWrite(); err != nil {
		return err
	}
	c, err := r.s.Client()
	if err != nil {
		return err
	}
	var user struct {
		ID string `json:"id"`
	}
	if err := c.Get(r.s.Ctx(), "/users/"+url.PathEscape(upn), url.Values{"$select": {"id"}}, &user); err != nil {
		return err
	}
	body := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + user.ID}
	err = c.Post(r.s.Ctx(), "/directoryRoles/"+url.PathEscape(roleID)+"/members/$ref", body, nil)
	r.s.Record("roles.addMember", roleID, "user="+upn, err)
	return err
}

// RemoveMember removes a user from a directory role. Destructive: typed confirm on UPN.
func (r *RolesService) RemoveMember(roleID, upn, confirm string) error {
	if err := r.s.GuardDestructive(upn, confirm); err != nil {
		return err
	}
	c, err := r.s.Client()
	if err != nil {
		return err
	}
	var user struct {
		ID string `json:"id"`
	}
	if err := c.Get(r.s.Ctx(), "/users/"+url.PathEscape(upn), url.Values{"$select": {"id"}}, &user); err != nil {
		return err
	}
	err = c.Delete(r.s.Ctx(), "/directoryRoles/"+url.PathEscape(roleID)+"/members/"+url.PathEscape(user.ID)+"/$ref")
	r.s.Record("roles.removeMember", roleID, "user="+upn, err)
	return err
}
