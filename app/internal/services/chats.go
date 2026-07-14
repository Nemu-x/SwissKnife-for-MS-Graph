package services

import (
	"encoding/json"
	"net/url"
	"strconv"
	"strings"

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

// ChatPickItem is a chat with a human label (topic, or the other participants).
type ChatPickItem struct {
	ID       string `json:"id"`
	Label    string `json:"label"`
	ChatType string `json:"chatType"`
}

// ListForPicker returns the user's chats labeled by topic or, for 1:1/group
// chats without a topic, by the other participants' names — so they're findable.
func (ch *ChatsService) ListForPicker(user string) ([]ChatPickItem, error) {
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{"$expand": {"members"}, "$top": {"50"}}
	raws, err := c.ListAll(ch.s.Ctx(), "/users/"+url.PathEscape(user)+"/chats", params, 0)
	if err != nil {
		return nil, err
	}
	userLower := strings.ToLower(user)
	out := make([]ChatPickItem, 0, len(raws))
	for _, raw := range raws {
		var chat struct {
			ID       string `json:"id"`
			Topic    string `json:"topic"`
			ChatType string `json:"chatType"`
			Members  []struct {
				DisplayName string `json:"displayName"`
				Email       string `json:"email"`
			} `json:"members"`
		}
		if json.Unmarshal(raw, &chat) != nil {
			continue
		}
		label := chat.Topic
		if label == "" {
			names := make([]string, 0, len(chat.Members))
			for _, m := range chat.Members {
				if strings.ToLower(m.Email) == userLower {
					continue
				}
				if m.DisplayName != "" {
					names = append(names, m.DisplayName)
				}
			}
			label = strings.Join(names, ", ")
		}
		if label == "" {
			label = chat.ID
		}
		out = append(out, ChatPickItem{ID: chat.ID, Label: label, ChatType: chat.ChatType})
	}
	return out, nil
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
