package actions

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/exoapi"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

type staticBroker struct{}

func (staticBroker) TokenFor(context.Context, string) (string, error) { return "exo", nil }

// exoCall is one Admin API request: cmdlet plus parameters.
type exoCall struct {
	Cmdlet string
	Params map[string]any
}

// exoHarness serves Graph and the Admin API from one fake server; exo
// answers each cmdlet.
func exoHarness(t *testing.T, graph func(w http.ResponseWriter, r *http.Request), exo func(c exoCall) string) (*engine.Engine, *[]exoCall) {
	t.Helper()
	var calls []exoCall
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if strings.HasPrefix(r.URL.Path, "/adminapi/v2.0/t1/") {
			var body struct {
				CmdletInput struct {
					CmdletName string
					Parameters map[string]any
				}
			}
			b, _ := io.ReadAll(r.Body)
			_ = json.Unmarshal(b, &body)
			c := exoCall{body.CmdletInput.CmdletName, body.CmdletInput.Parameters}
			calls = append(calls, c)
			w.Write([]byte(exo(c)))
			return
		}
		graph(w, r)
	}))
	t.Cleanup(srv.Close)
	old := exoapi.BaseURL
	exoapi.BaseURL = srv.URL
	t.Cleanup(func() { exoapi.BaseURL = old })

	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	s.SetTokens(staticBroker{})
	s.SetIdentity("t1", true)
	e := engine.New(s, engine.GraphProvider{}, exoapi.NewProvider())
	e.Register(Builtin()...)
	return e, &calls
}

func graphUsers(w http.ResponseWriter, r *http.Request) {
	switch {
	case r.URL.Path == "/organization":
		w.Write([]byte(`{"value":[{"verifiedDomains":[{"name":"contoso.com"},{"name":"contoso.onmicrosoft.com","isInitial":true}]}]}`))
	case strings.HasSuffix(r.URL.Path, "/calendar"):
		w.Write([]byte(`{"name":"Календарь"}`))
	case strings.HasPrefix(r.URL.Path, "/users/bob"):
		w.Write([]byte(`{"id":"b1","userPrincipalName":"bob@contoso.com","mail":"bob@contoso.com","displayName":"Bob Smith","mailNickname":"bob"}`))
	default:
		w.Write([]byte(`{"userPrincipalName":"ann@contoso.com","mail":"ann@contoso.com","displayName":"Ann"}`))
	}
}

func writesOf(calls []exoCall) []exoCall {
	var out []exoCall
	for _, c := range calls {
		if !strings.HasPrefix(c.Cmdlet, "Get-") {
			out = append(out, c)
		}
	}
	return out
}

func TestSendOnBehalfAddsIncrementally(t *testing.T) {
	e, calls := exoHarness(t, graphUsers, func(c exoCall) string {
		if c.Cmdlet == "Get-Mailbox" {
			return `{"value":[{"UserPrincipalName":"ann@contoso.com","GrantSendOnBehalfTo":["carol@contoso.com"]}]}`
		}
		return ``
	})
	p, err := e.Plan("mailbox.sendOnBehalf", engine.Inputs{"mailbox": "ann@contoso.com", "delegate": "bob@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "add" || c.After != "bob@contoso.com" {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writesOf(*calls)
	if len(w) != 1 || w[0].Cmdlet != "Set-Mailbox" {
		t.Fatalf("writes %+v", w)
	}
	delta, _ := w[0].Params["GrantSendOnBehalfTo"].(map[string]any)
	if delta["@odata.type"] != "#Exchange.GenericHashTable" || delta["add"].([]any)[0] != "bob@contoso.com" {
		t.Fatalf("delta %+v", delta)
	}
}

func TestSendOnBehalfAlreadyDelegateIsNoOp(t *testing.T) {
	e, _ := exoHarness(t, graphUsers, func(c exoCall) string {
		// Exchange lists delegates by recipient name: here the alias.
		return `[{"UserPrincipalName":"ann@contoso.com","GrantSendOnBehalfTo":["bob"]}]`
	})
	p, err := e.Plan("mailbox.sendOnBehalf", engine.Inputs{"mailbox": "ann@contoso.com", "delegate": "bob@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Op != "none" {
		t.Fatalf("change = %+v", p.Changes[0])
	}
}

func TestFolderPermissionUsesLocalizedCalendarAndSets(t *testing.T) {
	e, calls := exoHarness(t, graphUsers, func(c exoCall) string {
		if c.Cmdlet == "Get-MailboxFolderPermission" {
			return `{"value":[{"User":"Default","AccessRights":["AvailabilityOnly"]},{"User":"Bob Smith","AccessRights":["Reviewer"]}]}`
		}
		return ``
	})
	p, err := e.Plan("mailbox.folderPermission", engine.Inputs{"mailbox": "ann@contoso.com", "user": "bob@contoso.com", "access": "Editor"})
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "set" || c.Before != "Reviewer" || c.After != "Editor" || !strings.HasSuffix(c.Target, `ann@contoso.com:\Календарь`) {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writesOf(*calls)
	if len(w) != 1 || w[0].Cmdlet != "Set-MailboxFolderPermission" || w[0].Params["Identity"] != `ann@contoso.com:\Календарь` || w[0].Params["AccessRights"] != "Editor" {
		t.Fatalf("writes %+v", w)
	}
}

func TestFolderPermissionRejectsCalendarRoleOnInbox(t *testing.T) {
	e, _ := exoHarness(t, graphUsers, func(exoCall) string { return `` })
	if _, err := e.Plan("mailbox.folderPermission", engine.Inputs{"mailbox": "ann@contoso.com", "folder": "inbox", "user": "bob@contoso.com", "access": "LimitedDetails"}); err == nil {
		t.Fatal("calendar-only role on inbox must be rejected")
	}
}
