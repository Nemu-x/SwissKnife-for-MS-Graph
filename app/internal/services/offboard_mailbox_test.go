package services

import (
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// mailboxOffboardSession is harness() with the server URL exposed: the fake
// import URL must be absolute and point back at the fake.
func mailboxOffboardSession(t *testing.T, calls *[]string) *session.Session {
	t.Helper()
	var srvURL string
	srv := httptest.NewServer(mailboxOffboardHarness(t, calls, &srvURL))
	t.Cleanup(srv.Close)
	srvURL = srv.URL
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	return sess
}

// mailboxOffboardHarness extends the full offboard fake with the two mailboxes
// the mailbox copy step talks to (see mailboxtransfer_test.go for the shape).
func mailboxOffboardHarness(t *testing.T, calls *[]string, srvURL *string) http.HandlerFunc {
	t.Helper()
	base := offboardHarness(t, calls, map[string]string{})
	return func(w http.ResponseWriter, r *http.Request) {
		key := r.Method + " " + r.URL.Path
		switch key {
		case "GET /users/dep@contoso.com/settings/exchange":
			*calls = append(*calls, key)
			w.Write([]byte(`{"primaryMailboxId":"MBX:src@t"}`))
		case "GET /users/boss@contoso.com/settings/exchange":
			*calls = append(*calls, key)
			w.Write([]byte(`{"primaryMailboxId":"MBX:tgt@t"}`))
		case "GET /admin/exchange/mailboxes/MBX:src@t/folders":
			*calls = append(*calls, key)
			w.Write([]byte(`{"value":[{"id":"f-inbox","displayName":"Inbox","childFolderCount":0,"totalItemCount":1,"wellKnownName":"inbox","type":"IPF.Note"}]}`))
		case "GET /admin/exchange/mailboxes/MBX:src@t/folders/f-inbox/items":
			*calls = append(*calls, key)
			w.Write([]byte(`{"value":[{"id":"i1","size":10}]}`))
		case "POST /admin/exchange/mailboxes/MBX:tgt@t/folders":
			*calls = append(*calls, key)
			w.Write([]byte(`{"id":"t-root"}`))
		case "POST /admin/exchange/mailboxes/MBX:tgt@t/folders/t-root/childFolders":
			*calls = append(*calls, key)
			w.Write([]byte(`{"id":"t-inbox"}`))
		case "POST /admin/exchange/mailboxes/MBX:src@t/exportItems":
			*calls = append(*calls, key)
			w.Write([]byte(`{"value":[{"itemId":"i1","changeKey":"ck","data":"QUJD"}]}`))
		case "POST /admin/exchange/mailboxes/MBX:tgt@t/createImportSession":
			*calls = append(*calls, key)
			w.Write([]byte(`{"importUrl":"` + *srvURL + `/importItem?authtoken=x","expirationDateTime":"2099-01-01T00:00:00Z"}`))
		case "POST /importItem":
			*calls = append(*calls, key)
			w.Write([]byte(`{"itemId":"new-1"}`))
		default:
			base(w, r)
		}
	}
}

// The mailbox copy is a playbook step that must land after the OneDrive
// backup and before licenses are removed: without a license the mailbox is
// deleted, and with the account it is gone at once.
func TestOffboardCopiesMailboxBeforeLicenseRemoval(t *testing.T) {
	var calls []string
	sess := mailboxOffboardSession(t, &calls)

	req := fullOffboardRequest()
	req.BackupToUser = "archive@contoso.com"
	req.MailboxToUser = "boss@contoso.com"
	req.MailboxFolder = "Leaver archive"
	req.Delete = true
	res, err := NewPlaybookService(sess).Offboard(req)
	if err != nil {
		t.Fatal(err)
	}

	names := make([]string, 0, len(res.Steps))
	for _, s := range res.Steps {
		names = append(names, s.Name)
	}
	backup, mailbox, licenses, del := indexOf(names, "Backup OneDrive"), indexOf(names, "Copy mailbox"), indexOf(names, "Remove licenses"), indexOf(names, "Delete user")
	if backup < 0 || mailbox < 0 || licenses < 0 || del < 0 {
		t.Fatalf("expected backup, mailbox, licenses and delete steps, got %v", names)
	}
	if backup >= mailbox || mailbox >= licenses || licenses >= del {
		t.Errorf("order must be OneDrive backup → mailbox copy → licenses → delete, got %v", names)
	}
	st := res.Steps[mailbox]
	if !st.OK || st.NameKey != "steps.copyMailbox" || st.DetailKey != "stepDetails.mailbox" {
		t.Errorf("mailbox step wrong: %+v", st)
	}
	if st.Params["copied"] != 1 || st.Params["folders"] != 1 || st.Params["root"] != "Leaver archive" {
		t.Errorf("mailbox step params wrong: %v", st.Params)
	}
	if !strings.Contains(st.Detail, "1 item(s) in 1 folder(s)") {
		t.Errorf("mailbox step detail = %q", st.Detail)
	}
	// And on the wire: the export/import happened before the license call.
	exp, lic := indexOf(calls, "POST /admin/exchange/mailboxes/MBX:src@t/exportItems"), indexOf(calls, "POST /users/dep@contoso.com/assignLicense")
	imp := indexOf(calls, "POST /importItem")
	if exp < 0 || imp < 0 || lic < 0 || exp > lic || imp > lic {
		t.Errorf("mailbox export/import must precede license removal:\n%s", strings.Join(calls, "\n"))
	}
	if idx := indexOf(calls, "DELETE /users/dep@contoso.com"); idx >= 0 && idx < imp {
		t.Errorf("the account must outlive the mailbox copy:\n%s", strings.Join(calls, "\n"))
	}
}

// Without a mailbox target the step is skipped entirely — no step, no calls.
func TestOffboardSkipsMailboxCopyWithoutTarget(t *testing.T) {
	var calls []string
	sess := mailboxOffboardSession(t, &calls)
	res, err := NewPlaybookService(sess).Offboard(fullOffboardRequest())
	if err != nil {
		t.Fatal(err)
	}
	for _, s := range res.Steps {
		if s.Name == "Copy mailbox" {
			t.Errorf("mailbox step must not run without a target: %+v", s)
		}
	}
	for _, c := range calls {
		if strings.Contains(c, "/admin/exchange") || strings.Contains(c, "/settings/exchange") {
			t.Errorf("no mailbox API call without a target: %s", c)
		}
	}
}
