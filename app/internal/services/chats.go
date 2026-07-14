package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

type ChatsService struct {
	s *session.Session
}

func NewChatsService(s *session.Session) *ChatsService { return &ChatsService{s: s} }

func (ch *ChatsService) List(user string, maxItems int) ([]json.RawMessage, error) {
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(ch.s.Ctx(), "/users/"+url.PathEscape(user)+"/chats", topParams(50), maxItems)
}

func (ch *ChatsService) Messages(chatID string, top int) ([]json.RawMessage, error) {
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	if top <= 0 {
		top = 50
	}
	return c.ListAll(ch.s.Ctx(), "/chats/"+url.PathEscape(chatID)+"/messages", topParams(top), top)
}

func (ch *ChatsService) Members(chatID string) ([]json.RawMessage, error) {
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(ch.s.Ctx(), "/chats/"+url.PathEscape(chatID)+"/members", nil, 0)
}

func (ch *ChatsService) AddMember(chatID, upn string, asOwner bool) (json.RawMessage, error) {
	if err := ch.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Post(ch.s.Ctx(), "/chats/"+url.PathEscape(chatID)+"/members", conversationMember(upn, asOwner), &out)
	ch.s.Record("chats.addMember", chatID, "user="+upn, err)
	return out, err
}

func (ch *ChatsService) RemoveMember(chatID, upn string) error {
	if err := ch.s.GuardWrite(); err != nil {
		return err
	}
	c, err := ch.s.Client()
	if err != nil {
		return err
	}
	members, err := ch.Members(chatID)
	if err != nil {
		return err
	}
	mid, err := findMembershipID(members, upn)
	if err != nil {
		return err
	}
	err = c.Delete(ch.s.Ctx(), "/chats/"+url.PathEscape(chatID)+"/members/"+url.PathEscape(mid))
	ch.s.Record("chats.removeMember", chatID, "user="+upn, err)
	return err
}

// CreateGroupChat creates a group chat with a topic and members (at least 2 UPNs).
func (ch *ChatsService) CreateGroupChat(topic string, memberUpns []string) (json.RawMessage, error) {
	if err := ch.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	members := make([]map[string]any, 0, len(memberUpns))
	for _, upn := range memberUpns {
		members = append(members, map[string]any{
			"@odata.type":     "#microsoft.graph.aadUserConversationMember",
			"roles":           []string{"owner"},
			"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + upn + "')",
		})
	}
	body := map[string]any{
		"chatType": "group",
		"topic":    topic,
		"members":  members,
	}
	var out json.RawMessage
	err = c.Post(ch.s.Ctx(), "/chats", body, &out)
	ch.s.Record("chats.createGroup", topic, "members="+strconv.Itoa(len(memberUpns)), err)
	return out, err
}
