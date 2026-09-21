package services

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// mailboxFake is a two-mailbox Exchange fake for the import/export flow: the
// leaver (MBX:src@t) with a small folder tree, the manager (MBX:tgt@t) that
// receives the archive. It records every call and body for order assertions.
type mailboxFake struct {
	mu      sync.Mutex
	calls   []string
	bodies  map[string][]string // "METHOD path" → bodies, in order
	created int
	srvURL  string
}

func (f *mailboxFake) handler(t *testing.T) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		key := r.Method + " " + r.URL.Path
		body, _ := io.ReadAll(r.Body)
		f.mu.Lock()
		f.calls = append(f.calls, key)
		if len(body) > 0 {
			f.bodies[key] = append(f.bodies[key], string(body))
		}
		f.mu.Unlock()
		switch {
		case key == "GET /users/alice@contoso.com/settings/exchange":
			w.Write([]byte(`{"primaryMailboxId":"MBX:src@t"}`))
		case key == "GET /users/boss@contoso.com/settings/exchange":
			w.Write([]byte(`{"primaryMailboxId":"MBX:tgt@t"}`))
		case key == "GET /users/alice@contoso.com":
			w.Write([]byte(`{"displayName":"Alice Leaver"}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders":
			w.Write([]byte(`{"value":[
				{"id":"f-inbox","displayName":"Inbox","childFolderCount":1,"totalItemCount":2,"wellKnownName":"inbox","type":"IPF.Note"},
				{"id":"f-contacts","displayName":"Contacts","childFolderCount":0,"totalItemCount":1,"wellKnownName":"contacts","type":"IPF.Contact"},
				{"id":"f-cal","displayName":"Calendar","childFolderCount":0,"totalItemCount":1,"wellKnownName":"calendar","type":"IPF.Appointment"},
				{"id":"f-junk","displayName":"Junk Email","childFolderCount":0,"totalItemCount":9,"wellKnownName":"junkemail","type":"IPF.Note"},
				{"id":"f-tasks","displayName":"Tasks","childFolderCount":0,"totalItemCount":3,"wellKnownName":"tasks","type":"IPF.Task"}]}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-inbox/childFolders":
			w.Write([]byte(`{"value":[
				{"id":"f-proj","displayName":"Projects","childFolderCount":0,"totalItemCount":1,"wellKnownName":null,"type":"IPF.Note"}]}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-inbox/items":
			if got := r.URL.Query().Get("$select"); !strings.Contains(got, "size") {
				t.Errorf("item listing must select sizes for batching, got $select=%q", got)
			}
			w.Write([]byte(`{"value":[{"id":"i1","size":1000},{"id":"i2","size":2000}]}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-proj/items":
			w.Write([]byte(`{"value":[{"id":"i3","size":500}]}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-contacts/items":
			w.Write([]byte(`{"value":[{"id":"c1","size":300}]}`))
		case key == "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-cal/items":
			w.Write([]byte(`{"value":[{"id":"e1","size":400}]}`))
		case key == "POST /admin/exchange/mailboxes/MBX:tgt@t/folders",
			strings.HasPrefix(key, "POST /admin/exchange/mailboxes/MBX:tgt@t/folders/") && strings.HasSuffix(key, "/childFolders"):
			var req struct {
				DisplayName string `json:"displayName"`
				Type        string `json:"type"`
			}
			_ = json.Unmarshal(body, &req)
			if req.DisplayName == "" || req.Type == "" {
				t.Errorf("folder create needs displayName and type, got %s", body)
			}
			f.mu.Lock()
			f.created++
			id := "t-" + strings.ToLower(strings.ReplaceAll(req.DisplayName, " ", "-"))
			f.mu.Unlock()
			w.WriteHeader(http.StatusCreated)
			w.Write([]byte(`{"id":"` + id + `","displayName":"` + req.DisplayName + `","type":"` + req.Type + `"}`))
		case key == "POST /admin/exchange/mailboxes/MBX:src@t/exportItems":
			var req struct {
				ItemIDs []string `json:"itemIds"`
			}
			_ = json.Unmarshal(body, &req)
			if len(req.ItemIDs) == 0 || len(req.ItemIDs) > 20 {
				t.Errorf("exportItems takes 1..20 itemIds, got %d", len(req.ItemIDs))
			}
			parts := make([]string, 0, len(req.ItemIDs))
			for _, id := range req.ItemIDs {
				parts = append(parts, `{"itemId":"`+id+`","changeKey":"ck","data":"RlRTLSA`+id+`"}`)
			}
			w.Write([]byte(`{"value":[` + strings.Join(parts, ",") + `]}`))
		case key == "POST /admin/exchange/mailboxes/MBX:tgt@t/createImportSession":
			if len(body) > 0 {
				t.Errorf("createImportSession takes no body, got %s", body)
			}
			w.Write([]byte(`{"importUrl":"` + f.srvURL + `/api/gv1.0/Mailboxes('MBX:tgt@t')/importItem?authtoken=tok","expirationDateTime":"2099-01-01T00:00:00Z"}`))
		case key == "POST /api/gv1.0/Mailboxes('MBX:tgt@t')/importItem":
			if got := r.Header.Get("Authorization"); got != "" {
				t.Errorf("import URL is pre-authorized — no Bearer allowed, got %q", got)
			}
			if r.URL.Query().Get("authtoken") != "tok" {
				t.Errorf("import URL must be used verbatim, got %s", r.URL.String())
			}
			var req struct {
				FolderID string `json:"FolderId"`
				Mode     string `json:"Mode"`
				Data     string `json:"Data"`
			}
			_ = json.Unmarshal(body, &req)
			if req.Mode != "create" || req.FolderID == "" || !strings.HasPrefix(req.Data, "RlRTLSA") {
				t.Errorf("import body wrong: %s", body)
			}
			w.Write([]byte(`{"itemId":"new-` + strings.TrimPrefix(req.Data, "RlRTLSA") + `","changeKey":"ck2"}`))
		default:
			w.WriteHeader(http.StatusNotFound)
			w.Write([]byte(`{"error":{"code":"NotFound","message":"unexpected call ` + key + `"}}`))
		}
	}
}

func newMailboxHarness(t *testing.T) (*mailboxFake, *session.Session) {
	t.Helper()
	fake := &mailboxFake{bodies: map[string][]string{}}
	srv := httptest.NewServer(fake.handler(t))
	t.Cleanup(srv.Close)
	fake.srvURL = srv.URL
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	return fake, sess
}

func indexOf(calls []string, want string) int {
	for i, c := range calls {
		if c == want {
			return i
		}
	}
	return -1
}

// The preview walks the folder tree (root, then child folders of folders that
// have any), classifies by folder class and flags system folders — read-only:
// no items are listed, nothing is exported.
func TestMailboxPreviewClassifiesFolders(t *testing.T) {
	fake, sess := newMailboxHarness(t)
	pv, err := NewMailboxTransferService(sess).Preview("alice@contoso.com")
	if err != nil {
		t.Fatal(err)
	}
	if pv.MailboxID != "MBX:src@t" || pv.DisplayName != "Alice Leaver" {
		t.Errorf("identity wrong: %+v", pv)
	}
	// Inbox + Inbox/Projects are mail; Junk is a system folder even though it
	// is IPF.Note; Tasks is "other" and never copied.
	if pv.MailFolders != 2 || pv.MailItems != 3 || pv.ContactFolders != 1 || pv.ContactItems != 1 ||
		pv.CalendarFolders != 1 || pv.CalendarItems != 1 || pv.OtherFolders != 1 || pv.OtherItems != 3 || pv.SystemFolders != 1 {
		t.Errorf("counts wrong: %+v", pv)
	}
	var proj *MailboxFolderInfo
	for i := range pv.Folders {
		if pv.Folders[i].ID == "f-proj" {
			proj = &pv.Folders[i]
		}
	}
	if proj == nil || proj.Path != "Inbox/Projects" || proj.Parent != "Inbox" || proj.Kind != "mail" {
		t.Errorf("child folder path wrong: %+v", proj)
	}
	for _, c := range fake.calls {
		if strings.Contains(c, "/items") || strings.Contains(c, "exportItems") || strings.Contains(c, "MBX:tgt@t") {
			t.Errorf("preview must be read-only on the source: %s", c)
		}
	}
}

// Copy: resolve both mailbox ids, create the archive root and the folder tree
// in the target, export each folder's items in batches and import every
// stream into the matching target folder through the pre-authorized import
// URL. Contacts travel only when asked; junk and tasks never.
func TestMailboxCopyRecreatesTreeAndImportsItems(t *testing.T) {
	fake, sess := newMailboxHarness(t)
	res, err := NewMailboxTransferService(sess).Copy(MailboxCopyRequest{
		Source: "alice@contoso.com", Target: "boss@contoso.com", IncludeCalendar: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if res.RootFolder != "Archive - Alice Leaver" {
		t.Errorf("default root folder = %q", res.RootFolder)
	}
	// Inbox, Inbox/Projects, Calendar recreated (root not counted); 4 items copied.
	if res.Folders != 3 || res.Copied != 4 || res.TotalItems != 4 || len(res.Failed) != 0 || res.Canceled {
		t.Fatalf("result = %+v", res)
	}

	calls := fake.calls
	joined := strings.Join(calls, "\n")
	// Root folder created in the target mailbox root with the archive name.
	roots := fake.bodies["POST /admin/exchange/mailboxes/MBX:tgt@t/folders"]
	if len(roots) != 1 || !strings.Contains(roots[0], `"displayName":"Archive - Alice Leaver"`) || !strings.Contains(roots[0], `"type":"IPF.Note"`) {
		t.Errorf("root folder create wrong: %v", roots)
	}
	// Inbox and Calendar under the root; Projects under Inbox, with their classes.
	underRoot := fake.bodies["POST /admin/exchange/mailboxes/MBX:tgt@t/folders/t-archive---alice-leaver/childFolders"]
	if len(underRoot) != 2 || !strings.Contains(underRoot[0], `"displayName":"Inbox"`) ||
		!strings.Contains(underRoot[1], `"displayName":"Calendar"`) || !strings.Contains(underRoot[1], `"type":"IPF.Appointment"`) {
		t.Errorf("top-level folder creates wrong: %v", underRoot)
	}
	if got := fake.bodies["POST /admin/exchange/mailboxes/MBX:tgt@t/folders/t-inbox/childFolders"]; len(got) != 1 || !strings.Contains(got[0], `"displayName":"Projects"`) {
		t.Errorf("Projects must be created under the copied Inbox: %v", got)
	}
	// Excluded folders are neither created nor read.
	for _, bad := range []string{"f-contacts/items", "f-junk", "f-tasks", `"displayName":"Contacts"`, `"displayName":"Junk Email"`} {
		if strings.Contains(joined, bad) {
			t.Errorf("must not touch %s:\n%s", bad, joined)
		}
		for _, bodies := range fake.bodies {
			for _, b := range bodies {
				if strings.Contains(b, bad) {
					t.Errorf("must not create %s: %s", bad, b)
				}
			}
		}
	}
	// Export batches: one call per folder here (≤ 20 ids), source mailbox only.
	exports := fake.bodies["POST /admin/exchange/mailboxes/MBX:src@t/exportItems"]
	if len(exports) != 3 || !strings.Contains(exports[0], `"itemIds":["i1","i2"]`) {
		t.Errorf("export bodies wrong: %v", exports)
	}
	// One import session for the target, then one import per item, each into
	// the target folder that mirrors the item's source folder.
	if n := strings.Count(joined, "POST /admin/exchange/mailboxes/MBX:tgt@t/createImportSession"); n != 1 {
		t.Errorf("import session must be created once, got %d", n)
	}
	imports := fake.bodies["POST /api/gv1.0/Mailboxes('MBX:tgt@t')/importItem"]
	if len(imports) != 4 {
		t.Fatalf("want 4 imports, got %d: %v", len(imports), imports)
	}
	wantFolder := map[string]string{"i1": "t-inbox", "i2": "t-inbox", "i3": "t-projects", "e1": "t-calendar"}
	for _, b := range imports {
		var req struct {
			FolderID string `json:"FolderId"`
			Data     string `json:"Data"`
		}
		_ = json.Unmarshal([]byte(b), &req)
		id := strings.TrimPrefix(req.Data, "RlRTLSA")
		if wantFolder[id] != req.FolderID {
			t.Errorf("item %s imported into %s, want %s", id, req.FolderID, wantFolder[id])
		}
	}
	// Order: the folder exists before anything is exported for it, and the
	// export precedes its import.
	if indexOf(calls, "POST /admin/exchange/mailboxes/MBX:tgt@t/folders/t-archive---alice-leaver/childFolders") > indexOf(calls, "POST /admin/exchange/mailboxes/MBX:src@t/exportItems") {
		t.Errorf("target folders must exist before the first export:\n%s", joined)
	}
	if indexOf(calls, "POST /admin/exchange/mailboxes/MBX:src@t/exportItems") > indexOf(calls, "POST /api/gv1.0/Mailboxes('MBX:tgt@t')/importItem") {
		t.Errorf("export must precede import:\n%s", joined)
	}
}

// Opting contacts in copies the Contacts folder as an IPF.Contact folder; a
// custom root name is used verbatim.
func TestMailboxCopyHonorsContactsOptionAndFolderName(t *testing.T) {
	fake, sess := newMailboxHarness(t)
	res, err := NewMailboxTransferService(sess).Copy(MailboxCopyRequest{
		Source: "alice@contoso.com", Target: "boss@contoso.com", Folder: "Alice archive", IncludeContacts: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if res.RootFolder != "Alice archive" || res.Folders != 3 || res.Copied != 4 {
		t.Errorf("result = %+v", res)
	}
	under := strings.Join(fake.bodies["POST /admin/exchange/mailboxes/MBX:tgt@t/folders/t-alice-archive/childFolders"], "\n")
	if !strings.Contains(under, `"displayName":"Contacts","type":"IPF.Contact"`) {
		t.Errorf("contacts folder must be recreated with its class: %s", under)
	}
	if strings.Contains(under, "Calendar") {
		t.Errorf("calendar was not requested: %s", under)
	}
}

// A per-item export error is recorded and the copy carries on; a 403 aborts
// (every further item would fail the same way) and surfaces the Graph error.
func TestMailboxCopyRecordsItemFailuresAndAbortsOnForbidden(t *testing.T) {
	fake, sess := newMailboxHarness(t)
	base := fake.handler(t)
	forbid := false
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "POST" && strings.HasSuffix(r.URL.Path, "/exportItems") {
			body, _ := io.ReadAll(r.Body)
			if forbid {
				w.WriteHeader(http.StatusForbidden)
				w.Write([]byte(`{"error":{"code":"ErrorAccessDenied","message":"Access is denied"}}`))
				return
			}
			if strings.Contains(string(body), `"i3"`) {
				w.Write([]byte(`{"value":[{"itemId":"i3","error":{"code":"ErrorItemCorrupt","message":"corrupt"}}]}`))
				return
			}
			r.Body = io.NopCloser(strings.NewReader(string(body)))
		}
		base(w, r)
	}))
	t.Cleanup(srv.Close)
	fake.srvURL = srv.URL
	sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")

	res, err := NewMailboxTransferService(sess).Copy(MailboxCopyRequest{Source: "alice@contoso.com", Target: "boss@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if res.Copied != 2 || len(res.Failed) != 1 {
		t.Fatalf("want 2 copied + 1 failed, got %+v", res)
	}
	for k, v := range res.Failed {
		if !strings.HasPrefix(k, "Inbox/Projects · i3") || !strings.Contains(v, "ErrorItemCorrupt") {
			t.Errorf("failure entry wrong: %q → %q", k, v)
		}
	}

	forbid = true
	_, err = NewMailboxTransferService(sess).Copy(MailboxCopyRequest{Source: "alice@contoso.com", Target: "boss@contoso.com"})
	if err == nil || !strings.Contains(err.Error(), "ErrorAccessDenied") {
		t.Fatalf("a 403 must abort the copy with the Graph error, got %v", err)
	}
}

// Source and target must differ, and the copy is a write (read-only mode
// blocks it before any call).
func TestMailboxCopyGuards(t *testing.T) {
	fake, sess := newMailboxHarness(t)
	svc := NewMailboxTransferService(sess)
	if _, err := svc.Copy(MailboxCopyRequest{Source: "alice@contoso.com", Target: "Alice@contoso.com"}); err == nil {
		t.Error("same source and target must be rejected")
	}
	sess.SetReadOnly(true)
	if _, err := svc.Copy(MailboxCopyRequest{Source: "alice@contoso.com", Target: "boss@contoso.com"}); err == nil {
		t.Error("read-only mode must block the copy")
	}
	if len(fake.calls) != 0 {
		t.Errorf("guards must fire before any Graph call: %v", fake.calls)
	}
}

// Batching: 20 ids per export call at most, and the byte budget splits a run
// of large items so one response never carries gigabytes of base64.
func TestSplitExportBatches(t *testing.T) {
	var items []mbxItem
	for i := 0; i < 45; i++ {
		items = append(items, mbxItem{ID: itoa(i), Size: 1000})
	}
	b := splitExportBatches(items)
	if len(b) != 3 || len(b[0]) != 20 || len(b[1]) != 20 || len(b[2]) != 5 {
		t.Errorf("count split wrong: %d batches", len(b))
	}
	big := []mbxItem{{ID: "a", Size: 20 << 20}, {ID: "b", Size: 10 << 20}, {ID: "c", Size: 100 << 20}, {ID: "d", Size: 1}}
	b = splitExportBatches(big)
	// a+b would exceed the cap, c is over it on its own, and d must not ride
	// along with c — every batch stays under the cap unless a single item is.
	if len(b) != 4 || len(b[0]) != 1 || len(b[1]) != 1 || len(b[2]) != 1 || len(b[3]) != 1 {
		t.Errorf("size split wrong: %v", b)
	}
	if len(splitExportBatches(nil)) != 0 {
		t.Error("no items, no batches")
	}
}
