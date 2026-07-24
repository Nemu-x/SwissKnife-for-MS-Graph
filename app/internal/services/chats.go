package services

import (
	"context"
	"encoding/json"
	"fmt"
	"net/url"
	"os"
	"strconv"
	"strings"
	"time"

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

// ChatBackupResult summarizes an exported chat archive.
type ChatBackupResult struct {
	Chats    int    `json:"chats"`
	Messages int    `json:"messages"`
	File     string `json:"file"` // path uploaded into the target drive
}

// BackupUserChats exports a user's Teams chats (1:1 and group) into a single
// JSON archive and uploads it to the target user's OneDrive folder.
//
// NOTE: reading another user's chat messages app-only goes through Microsoft's
// PROTECTED APIs — Chat.Read.All (Application) must be approved by Microsoft
// for this app id (aka.ms/teamsgraph/requestaccess) before Graph stops
// answering 403. The permission hint in the UI points this out.
func (ch *ChatsService) BackupUserChats(sourceUser, targetUser, destFolder string) (*ChatBackupResult, error) {
	res, err := ch.backupUserChatsCtx(ch.s.Ctx(), sourceUser, targetUser, destFolder)
	return res, wrapOpErr(err)
}

func (ch *ChatsService) backupUserChatsCtx(parent context.Context, sourceUser, targetUser, destFolder string) (*ChatBackupResult, error) {
	if err := ch.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := ch.s.Client()
	if err != nil {
		return nil, err
	}

	// One paged walk over every chat the user participates in.
	raws, err := c.ListAll(parent, "/users/"+url.PathEscape(sourceUser)+"/chats/getAllMessages", topParams(50), 0)
	if err != nil {
		return nil, err
	}

	// Chat labels (topic / other participants) — best-effort; ids otherwise.
	labels := map[string]string{}
	if picks, lerr := ch.ListForPicker(sourceUser); lerr == nil {
		for _, p := range picks {
			labels[p.ID] = p.Label
		}
	}

	type chatMsg struct {
		From string `json:"from,omitempty"`
		At   string `json:"at,omitempty"`
		Type string `json:"contentType,omitempty"`
		Body string `json:"body,omitempty"`
	}
	grouped := map[string][]chatMsg{}
	order := []string{}
	for _, raw := range raws {
		var m struct {
			ChatID string `json:"chatId"`
			At     string `json:"createdDateTime"`
			From   struct {
				User struct {
					DisplayName string `json:"displayName"`
				} `json:"user"`
			} `json:"from"`
			Body struct {
				ContentType string `json:"contentType"`
				Content     string `json:"content"`
			} `json:"body"`
			MessageType string `json:"messageType"`
		}
		if json.Unmarshal(raw, &m) != nil || m.ChatID == "" {
			continue
		}
		if m.MessageType != "" && m.MessageType != "message" {
			continue // skip system events (member added, …)
		}
		if _, seen := grouped[m.ChatID]; !seen {
			order = append(order, m.ChatID)
		}
		grouped[m.ChatID] = append(grouped[m.ChatID], chatMsg{
			From: m.From.User.DisplayName, At: m.At, Type: m.Body.ContentType, Body: m.Body.Content,
		})
	}

	type chatOut struct {
		ID       string    `json:"id"`
		Label    string    `json:"label,omitempty"`
		Messages []chatMsg `json:"messages"`
	}
	res := &ChatBackupResult{Chats: len(order)}
	archive := struct {
		User       string    `json:"user"`
		ExportedAt string    `json:"exportedAt"`
		Chats      []chatOut `json:"chats"`
	}{User: sourceUser, ExportedAt: time.Now().Format(time.RFC3339)}
	for _, id := range order {
		archive.Chats = append(archive.Chats, chatOut{ID: id, Label: labels[id], Messages: grouped[id]})
		res.Messages += len(grouped[id])
	}

	blob, err := json.MarshalIndent(archive, "", "  ")
	if err != nil {
		return nil, err
	}
	tmp, err := os.CreateTemp("", "swissknife-chats-*.json")
	if err != nil {
		return nil, err
	}
	local := tmp.Name()
	defer os.Remove(local)
	if _, err := tmp.Write(blob); err != nil {
		tmp.Close()
		return nil, err
	}
	tmp.Close()

	remote := strings.Trim(destFolder, "/")
	name := fmt.Sprintf("teams-chats-%s.json", time.Now().Format("2006-01-02-1504"))
	if remote != "" {
		remote += "/"
	}
	dst := "/users/" + url.PathEscape(targetUser) + "/drive/root:/" + escapeDrivePath(remote+name)
	if _, err := c.UploadFile(parent, dst, local, nil); err != nil {
		return nil, err
	}
	res.File = remote + name

	ch.s.Record("chats.backup", sourceUser+" -> "+targetUser,
		"chats="+itoa(res.Chats)+" messages="+itoa(res.Messages)+" file="+res.File, nil)
	return res, nil
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
