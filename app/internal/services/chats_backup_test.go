package services

import (
	"io"
	"net/http"
	"strings"
	"testing"
)

// TestBackupUserChatsExportsAndUploads: messages from getAllMessages are
// grouped per chat, system events skipped, and one JSON archive is uploaded
// into the target user's OneDrive folder.
func TestBackupUserChatsExportsAndUploads(t *testing.T) {
	var uploadPath, uploadBody string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/chats/getAllMessages"):
			w.Write([]byte(`{"value":[
				{"chatId":"c1","createdDateTime":"2026-07-25T10:00:00Z","messageType":"message",
				 "from":{"user":{"displayName":"Alice"}},"body":{"contentType":"text","content":"hi"}},
				{"chatId":"c1","createdDateTime":"2026-07-25T10:01:00Z","messageType":"message",
				 "from":{"user":{"displayName":"Bob"}},"body":{"contentType":"text","content":"hello"}},
				{"chatId":"c2","createdDateTime":"2026-07-25T11:00:00Z","messageType":"message",
				 "from":{"user":{"displayName":"Alice"}},"body":{"contentType":"html","content":"<p>yo</p>"}},
				{"chatId":"c2","createdDateTime":"2026-07-25T11:01:00Z","messageType":"systemEventMessage",
				 "body":{"contentType":"html","content":"member added"}}]}`))
		case r.Method == "PUT" && strings.Contains(r.URL.Path, ":/content"):
			uploadPath = r.URL.Path
			b, _ := io.ReadAll(r.Body)
			uploadBody = string(b)
			w.Write([]byte(`{"id":"file1"}`))
		default:
			w.Write([]byte(`{}`))
		}
	})
	chats := NewChatsService(sess)

	res, err := chats.BackupUserChats("dep@contoso.com", "arch@contoso.com", "dep@contoso.com")
	if err != nil {
		t.Fatal(err)
	}
	if res.Chats != 2 || res.Messages != 3 {
		t.Fatalf("want 2 chats / 3 messages (system event skipped), got %+v", res)
	}
	if !strings.Contains(uploadPath, "/users/arch@contoso.com/drive/root:/dep@contoso.com/teams-chats-") {
		t.Fatalf("upload path wrong: %s", uploadPath)
	}
	if !strings.Contains(uploadBody, `"hi"`) || !strings.Contains(uploadBody, "Alice") {
		t.Fatalf("archive body missing messages: %.200s", uploadBody)
	}
	if strings.Contains(uploadBody, "member added") {
		t.Fatal("system events must be excluded")
	}
}

// TestOffboardChatsStepUsesHint: a 403 on the chats export surfaces the
// protected-API hint so the operator knows this permission needs Microsoft
// approval, not just admin consent.
func TestOffboardChatsStepUsesHint(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		if strings.HasSuffix(r.URL.Path, "/chats/getAllMessages") {
			w.WriteHeader(403)
			w.Write([]byte(`{"error":{"code":"Forbidden","message":"Access denied."}}`))
			return
		}
		w.Write([]byte(`{}`))
	})
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.BackupToUser = "archive@contoso.com"
	req.BackupChats = true
	res, err := pb.Offboard(req)
	if err != nil {
		t.Fatal(err)
	}
	for _, st := range res.Steps {
		if st.Name == "Backup Teams chats" {
			if st.OK {
				t.Fatalf("chats step must fail on 403: %+v", st)
			}
			if !strings.Contains(st.Hint, "Chat.Read.All") {
				t.Fatalf("hint must name the protected permission, got %q", st.Hint)
			}
			return
		}
	}
	t.Fatal("Backup Teams chats step missing")
}
