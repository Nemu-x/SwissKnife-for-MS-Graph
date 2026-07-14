package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

type GroupsService struct {
	s *session.Session
}

func NewGroupsService(s *session.Session) *GroupsService { return &GroupsService{s: s} }

func (g *GroupsService) List(search string, maxItems int) ([]json.RawMessage, error) {
	c, err := g.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{"$top": {"100"}}
	if search != "" {
		params.Set("$filter", "startswith(displayName,'"+escapeODataLiteral(search)+"')")
	}
	return c.ListAll(g.s.Ctx(), "/groups", params, maxItems)
}

func (g *GroupsService) Get(groupID string) (json.RawMessage, error) {
	c, err := g.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(g.s.Ctx(), "/groups/"+url.PathEscape(groupID), nil, &out)
	return out, err
}

func (g *GroupsService) Members(groupID string) ([]json.RawMessage, error) {
	c, err := g.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(g.s.Ctx(), "/groups/"+url.PathEscape(groupID)+"/members", nil, 0)
}

func (g *GroupsService) addRef(groupID, upn, kind string) error {
	if err := g.s.GuardWrite(); err != nil {
		return err
	}
	c, err := g.s.Client()
	if err != nil {
		return err
	}
	// first resolve the UPN into an id
	var user struct {
		ID string `json:"id"`
	}
	if err := c.Get(g.s.Ctx(), "/users/"+url.PathEscape(upn), url.Values{"$select": {"id"}}, &user); err != nil {
		return err
	}
	body := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/users/" + user.ID}
	err = c.Post(g.s.Ctx(), "/groups/"+url.PathEscape(groupID)+"/"+kind+"/$ref", body, nil)
	g.s.Record("groups.add_"+kind, groupID, "user="+upn, err)
	return err
}

func (g *GroupsService) AddOwner(groupID, upn string) error {
	return g.addRef(groupID, upn, "owners")
}

func (g *GroupsService) AddMember(groupID, upn string) error {
	return g.addRef(groupID, upn, "members")
}

// CreateM365 creates a Unified group; ownerUpn is optional.
func (g *GroupsService) CreateM365(displayName, description, mailNickname, ownerUpn string) (json.RawMessage, error) {
	if err := g.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := g.s.Client()
	if err != nil {
		return nil, err
	}
	body := map[string]any{
		"displayName":     displayName,
		"description":     description,
		"groupTypes":      []string{"Unified"},
		"mailEnabled":     true,
		"securityEnabled": false,
		"mailNickname":    mailNickname,
	}
	if ownerUpn != "" {
		userURL := "https://graph.microsoft.com/v1.0/users('" + ownerUpn + "')"
		body["owners@odata.bind"] = []string{userURL}
		body["members@odata.bind"] = []string{userURL}
	}
	var out json.RawMessage
	err = c.Post(g.s.Ctx(), "/groups", body, &out)
	g.s.Record("groups.create", displayName, "nickname="+mailNickname, err)
	return out, err
}
