package services

import (
	"encoding/json"
	"errors"
	"net/url"
	"strings"

	"swissknife-app/internal/session"
)

type TeamsService struct {
	s *session.Session
}

func NewTeamsService(s *session.Session) *TeamsService { return &TeamsService{s: s} }

func (t *TeamsService) JoinedTeams(user string) ([]json.RawMessage, error) {
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(t.s.Ctx(), "/users/"+url.PathEscape(user)+"/joinedTeams", nil, 0)
}

func (t *TeamsService) Channels(teamID string) ([]json.RawMessage, error) {
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/channels", nil, 0)
}

func (t *TeamsService) TeamMembers(teamID string) ([]json.RawMessage, error) {
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/members", nil, 0)
}

func (t *TeamsService) ChannelMembers(teamID, channelID string) ([]json.RawMessage, error) {
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/channels/"+url.PathEscape(channelID)+"/members", nil, 0)
}

func conversationMember(upn string, asOwner bool) map[string]any {
	roles := []string{}
	if asOwner {
		roles = []string{"owner"}
	}
	return map[string]any{
		"@odata.type":     "#microsoft.graph.aadUserConversationMember",
		"roles":           roles,
		"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + upn + "')",
	}
}

func (t *TeamsService) AddTeamMember(teamID, upn string, asOwner bool) (json.RawMessage, error) {
	if err := t.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Post(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/members", conversationMember(upn, asOwner), &out)
	t.s.Record("teams.addMember", teamID, "user="+upn, err)
	return out, err
}

func (t *TeamsService) AddChannelMember(teamID, channelID, upn string, asOwner bool) (json.RawMessage, error) {
	if err := t.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Post(t.s.Ctx(),
		"/teams/"+url.PathEscape(teamID)+"/channels/"+url.PathEscape(channelID)+"/members",
		conversationMember(upn, asOwner), &out)
	t.s.Record("teams.addChannelMember", teamID+"/"+channelID, "user="+upn, err)
	return out, err
}

// findMembershipID looks up a membership id by email in the member list.
func findMembershipID(items []json.RawMessage, upn string) (string, error) {
	target := strings.ToLower(upn)
	for _, raw := range items {
		var m struct {
			ID    string `json:"id"`
			Email string `json:"email"`
		}
		if json.Unmarshal(raw, &m) == nil && strings.ToLower(m.Email) == target {
			return m.ID, nil
		}
	}
	return "", errors.New("member with UPN/email " + upn + " not found")
}

func (t *TeamsService) RemoveTeamMember(teamID, upn string) error {
	if err := t.s.GuardWrite(); err != nil {
		return err
	}
	c, err := t.s.Client()
	if err != nil {
		return err
	}
	members, err := t.TeamMembers(teamID)
	if err != nil {
		return err
	}
	mid, err := findMembershipID(members, upn)
	if err != nil {
		return err
	}
	err = c.Delete(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/members/"+url.PathEscape(mid))
	t.s.Record("teams.removeMember", teamID, "user="+upn, err)
	return err
}

func (t *TeamsService) RemoveChannelMember(teamID, channelID, upn string) error {
	if err := t.s.GuardWrite(); err != nil {
		return err
	}
	c, err := t.s.Client()
	if err != nil {
		return err
	}
	members, err := t.ChannelMembers(teamID, channelID)
	if err != nil {
		return err
	}
	mid, err := findMembershipID(members, upn)
	if err != nil {
		return err
	}
	err = c.Delete(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/channels/"+url.PathEscape(channelID)+"/members/"+url.PathEscape(mid))
	t.s.Record("teams.removeChannelMember", teamID+"/"+channelID, "user="+upn, err)
	return err
}

// Teamify converts an M365 group into a Team.
func (t *TeamsService) Teamify(groupID string) (json.RawMessage, error) {
	if err := t.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	body := map[string]any{
		"memberSettings":    map[string]any{"allowCreateUpdateChannels": true},
		"messagingSettings": map[string]any{"allowUserEditMessages": true, "allowUserDeleteMessages": true},
		"funSettings":       map[string]any{"allowGiphy": true, "giphyContentRating": "strict"},
	}
	var out json.RawMessage
	err = c.Put(t.s.Ctx(), "/groups/"+url.PathEscape(groupID)+"/team", body, &out)
	t.s.Record("teams.teamify", groupID, "", err)
	return out, err
}

// CreateChannel: standard | private | shared; private/shared require ownerUpn.
func (t *TeamsService) CreateChannel(teamID, displayName, description, channelType, ownerUpn string) (json.RawMessage, error) {
	if err := t.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := t.s.Client()
	if err != nil {
		return nil, err
	}
	body := map[string]any{
		"displayName":    displayName,
		"description":    description,
		"membershipType": channelType,
	}
	if channelType == "private" || channelType == "shared" {
		if ownerUpn == "" {
			return nil, errors.New("private/shared channel requires an owner UPN")
		}
		body["members"] = []map[string]any{conversationMember(ownerUpn, true)}
	}
	var out json.RawMessage
	err = c.Post(t.s.Ctx(), "/teams/"+url.PathEscape(teamID)+"/channels", body, &out)
	t.s.Record("teams.createChannel", teamID, "name="+displayName+" type="+channelType, err)
	return out, err
}
