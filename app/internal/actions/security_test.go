package actions

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

func securityHarness(t *testing.T, h http.HandlerFunc, fake *fakePS) *engine.Engine {
	t.Helper()
	srv := httptest.NewServer(h)
	t.Cleanup(srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{}, stubProvider{backend: pwsh.BackendExchangePS})
	if fake != nil {
		e.PS = fake
	}
	e.Register(Builtin()...)
	return e
}

func TestRuleAuditFlagsAttackerPatternsOnly(t *testing.T) {
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		p := r.URL.Path
		switch {
		case p == "/domains":
			w.Write([]byte(`{"value":[{"id":"contoso.com"}]}`))
		case p == "/users":
			w.Write([]byte(`{"value":[{"userPrincipalName":"ann@contoso.com","mail":"ann@contoso.com"},
				{"userPrincipalName":"bob@contoso.com","mail":"bob@contoso.com"},
				{"userPrincipalName":"svc@contoso.com"}]}`))
		case strings.HasSuffix(p, "/mailFolders/deleteditems"):
			w.Write([]byte(`{"id":"DEL"}`))
		case strings.HasSuffix(p, "/mailFolders/junkemail"), strings.HasSuffix(p, "/mailFolders/archive"), strings.HasSuffix(p, "/mailFolders/conversationhistory"):
			w.Write([]byte(`{"id":"X-` + p[len(p)-4:] + `"}`))
		case p == "/users/ann@contoso.com/mailFolders/inbox/messageRules":
			w.Write([]byte(`{"value":[
				{"displayName":"internal fwd","isEnabled":true,"actions":{"forwardTo":[{"emailAddress":{"address":"boss@contoso.com"}}]}},
				{"displayName":".","isEnabled":true,"actions":{"forwardTo":[{"emailAddress":{"address":"drop@evil.example"}}]}},
				{"displayName":"hide","isEnabled":true,"actions":{"moveToFolder":"DEL","markAsRead":true}},
				{"displayName":"file newsletters","isEnabled":true,"actions":{"moveToFolder":"NEWS","markAsRead":true}}]}`))
		case p == "/users/bob@contoso.com/mailFolders/inbox/messageRules":
			w.WriteHeader(http.StatusNotFound)
			w.Write([]byte(`{"error":{"code":"MailboxNotEnabledForRESTAPI","message":"no mailbox"}}`))
		default:
			t.Errorf("unexpected %s", p)
		}
	}, nil)
	res, err := e.Run(t.Context(), "mail.ruleAudit", engine.Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != 2 {
		t.Fatalf("rows %+v (want the external forward and the hide rule only)", res.Rows)
	}
	if res.Rows[0]["why"] != "externalForward=drop@evil.example" || res.Rows[1]["why"] != "hides" {
		t.Fatalf("rows %+v", res.Rows)
	}
	if res.Note == nil || res.Note.Params["scanned"] != "1" || res.Note.Params["unreadable"] != "1" {
		t.Fatalf("note %+v", res.Note)
	}
}

func TestReportThreatSubmitsTheNewestMessage(t *testing.T) {
	var submitted map[string]any
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users/ann@contoso.com":
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com"}`))
		case "/users/u1/messages":
			if !strings.Contains(r.URL.Query().Get("$filter"), "'o''brien@phish.example'") {
				t.Errorf("filter %q", r.URL.Query().Get("$filter"))
			}
			w.Write([]byte(`{"value":[{"id":"m1","subject":"Pay now","receivedDateTime":"2026-10-07T10:00:00Z"}]}`))
		case "/security/threatSubmission/emailThreats":
			b, _ := io.ReadAll(r.Body)
			_ = json.Unmarshal(b, &submitted)
			w.Write([]byte(`{}`))
		}
	}, nil)
	ch := planApply(t, e, "mail.reportThreat", engine.Inputs{"user": "ann@contoso.com", "sender": "o'brien@phish.example"})
	if !strings.Contains(ch.Target, "Pay now") {
		t.Fatalf("change %+v", ch)
	}
	if submitted["category"] != "phishing" || submitted["messageUrl"] != "https://graph.microsoft.com/v1.0/users/u1/messages/m1" {
		t.Fatalf("submitted %+v", submitted)
	}
}

func TestBlockSenderAndReleaseQuarantine(t *testing.T) {
	fake := &fakePS{answers: map[string]string{
		"Get-TenantAllowBlockListItems": `[]`,
		"Get-QuarantineMessage":         `[{"Subject":"Invoice","SenderAddress":"x@vendor.example","ReleaseStatus":"NotReleased"}]`,
	}}
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) }, fake)
	if _, err := e.Plan("mail.blockSender", engine.Inputs{"sender": "*.example"}); err == nil {
		t.Fatal("a wildcard entry must be refused")
	}
	ch := planApply(t, e, "mail.blockSender", engine.Inputs{"sender": "Phish.Example"})
	if ch.Op != "add" || ch.After != "phish.example" {
		t.Fatalf("change %+v", ch)
	}
	ch = planApply(t, e, "mail.releaseQuarantine", engine.Inputs{"identity": "abc\\def"})
	if ch.Before != "NotReleased" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if len(w) != 2 || w[0].Cmdlet != "New-TenantAllowBlockListItems" || w[0].Params["Block"] != true ||
		w[1].Cmdlet != "Release-QuarantineMessage" || w[1].Params["Identity"] != "abc\\def" {
		t.Fatalf("writes %+v", w)
	}
}
