package actions

import (
	"encoding/json"
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

// fakePS answers cmdlets from a table and records every call.
type fakePS struct {
	answers map[string]string // cmdlet → JSON array
	calls   []exoCall
}

func (f *fakePS) Invoke(_ engine.Env, family, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	if family != pwsh.FamilyExchange {
		panic("unexpected family " + family)
	}
	f.calls = append(f.calls, exoCall{cmdlet, params})
	var out []json.RawMessage
	if a, ok := f.answers[cmdlet]; ok {
		_ = json.Unmarshal([]byte(a), &out)
	}
	return out, nil
}

func (f *fakePS) writes() []exoCall {
	var out []exoCall
	for _, c := range f.calls {
		if !strings.HasPrefix(c.Cmdlet, "Get-") {
			out = append(out, c)
		}
	}
	return out
}

type stubProvider struct {
	backend engine.Backend
	reason  *engine.Reason
}

func (p stubProvider) Backend() engine.Backend                { return p.backend }
func (p stubProvider) Status(*session.Session) *engine.Reason { return p.reason }

// psHarness: Graph from graphUsers (plus groups), PowerShell from fake, the
// Admin API marked unavailable so its actions fall back to PowerShell.
func psHarness(t *testing.T, fake *fakePS) *engine.Engine {
	t.Helper()
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/groups/dl1":
			w.Write([]byte(`{"mail":"sales@contoso.com","displayName":"Sales","mailEnabled":true,"groupTypes":[]}`))
		case "/groups/m365":
			w.Write([]byte(`{"mail":"team@contoso.com","displayName":"Team","mailEnabled":true,"groupTypes":["Unified"]}`))
		default:
			graphUsers(w, r)
		}
	}))
	t.Cleanup(srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{},
		stubProvider{backend: "exo-api", reason: &engine.Reason{Key: "exoApiNotEnabled"}},
		stubProvider{backend: pwsh.BackendExchangePS})
	e.PS = fake
	e.Register(Builtin()...)
	return e
}

func planApply(t *testing.T, e *engine.Engine, id string, in engine.Inputs) engine.Change {
	t.Helper()
	p, err := e.Plan(id, in)
	if err != nil {
		t.Fatalf("plan %s: %v", id, err)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatalf("apply %s: %v", id, err)
	}
	return p.Changes[0]
}

func TestFullAccessGrantWithAutomapping(t *testing.T) {
	fake := &fakePS{answers: map[string]string{
		// An inherited or denied FullAccess does not count as granted.
		"Get-MailboxPermission": `[{"AccessRights":["FullAccess"],"IsInherited":true},{"AccessRights":["FullAccess"],"Deny":true}]`,
	}}
	ch := planApply(t, psHarness(t, fake), "mailbox.fullAccess", engine.Inputs{"mailbox": "ann@contoso.com", "user": "bob@contoso.com"})
	if ch.Op != "add" || ch.After != "bob@contoso.com" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if len(w) != 1 || w[0].Cmdlet != "Add-MailboxPermission" || w[0].Params["AccessRights"] != "FullAccess" || w[0].Params["AutoMapping"] != true {
		t.Fatalf("writes %+v", w)
	}
}

func TestMailboxTypeConvertsToShared(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-Mailbox": `[{"RecipientTypeDetails":"UserMailbox"}]`}}
	ch := planApply(t, psHarness(t, fake), "mailbox.type", engine.Inputs{"mailbox": "ann@contoso.com"})
	if ch.Before != "regular" || ch.After != "shared" {
		t.Fatalf("change %+v", ch)
	}
	if w := fake.writes(); len(w) != 1 || w[0].Params["Type"] != "Shared" {
		t.Fatalf("writes %+v", w)
	}

	fake = &fakePS{answers: map[string]string{"Get-Mailbox": `[{"RecipientTypeDetails":"RoomMailbox"}]`}}
	if _, err := psHarness(t, fake).Plan("mailbox.type", engine.Inputs{"mailbox": "ann@contoso.com"}); err == nil {
		t.Fatal("a room mailbox must not be converted")
	}
}

func TestForwardingSetAndClear(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-Mailbox": `[{"ForwardingSmtpAddress":null,"DeliverToMailboxAndForward":false}]`}}
	ch := planApply(t, psHarness(t, fake), "mailbox.forwarding", engine.Inputs{"mailbox": "ann@contoso.com", "forwardTo": "bob@contoso.com"})
	if ch.Op != "set" || ch.After != "bob@contoso.com" || ch.Note != "forwardKeepCopy" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if len(w) != 1 || w[0].Params["ForwardingSmtpAddress"] != "smtp:bob@contoso.com" || w[0].Params["DeliverToMailboxAndForward"] != true {
		t.Fatalf("writes %+v", w)
	}

	fake = &fakePS{answers: map[string]string{"Get-Mailbox": `[{"ForwardingSmtpAddress":"smtp:evil@external.com","DeliverToMailboxAndForward":false}]`}}
	ch = planApply(t, psHarness(t, fake), "mailbox.forwarding", engine.Inputs{"mailbox": "ann@contoso.com", "op": "clear"})
	if ch.Op != "remove" || ch.Before != "evil@external.com" {
		t.Fatalf("clear change %+v", ch)
	}
	w = fake.writes()
	if v, ok := w[0].Params["ForwardingSmtpAddress"]; !ok || v != nil {
		t.Fatalf("clear must send null: %+v", w)
	}

	if _, err := psHarness(t, &fakePS{answers: map[string]string{"Get-Mailbox": `[{}]`}}).Plan("mailbox.forwarding",
		engine.Inputs{"mailbox": "ann@contoso.com"}); err == nil {
		t.Fatal("set without a target must be refused")
	}
}

func TestAddressesNeverDropThePrimary(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-Mailbox": `[{"EmailAddresses":["SMTP:ann@contoso.com","smtp:a@contoso.com"]}]`}}
	e := psHarness(t, fake)
	if _, err := e.Plan("mailbox.address", engine.Inputs{"mailbox": "ann@contoso.com", "address": "ann@contoso.com", "op": "remove"}); err == nil {
		t.Fatal("removing the primary address must be refused")
	}
	ch := planApply(t, e, "mailbox.address", engine.Inputs{"mailbox": "ann@contoso.com", "address": "sales@contoso.com"})
	if ch.Op != "add" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if hash, _ := w[0].Params["EmailAddresses"].(map[string]any); hash["Add"] != "smtp:sales@contoso.com" {
		t.Fatalf("writes %+v", w)
	}
	if _, err := e.Plan("mailbox.address", engine.Inputs{"mailbox": "ann@contoso.com", "address": "not an address"}); err == nil {
		t.Fatal("a malformed address must be refused")
	}
}

func TestDistributionListMembership(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-DistributionGroupMember": `[{"PrimarySmtpAddress":"carol@contoso.com"}]`}}
	e := psHarness(t, fake)
	if _, err := e.Plan("distributionList.membership", engine.Inputs{"user": "bob@contoso.com", "group": "m365"}); err == nil {
		t.Fatal("an M365 group is not a distribution list")
	}
	ch := planApply(t, e, "distributionList.membership", engine.Inputs{"user": "bob@contoso.com", "group": "dl1"})
	if ch.Op != "add" || ch.After != "Sales" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if len(w) != 1 || w[0].Cmdlet != "Add-DistributionGroupMember" || w[0].Params["Identity"] != "sales@contoso.com" || w[0].Params["Member"] != "bob@contoso.com" {
		t.Fatalf("writes %+v", w)
	}
}

func TestTransportRuleDisable(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-TransportRule": `[{"Name":"Block exe","State":"Enabled"}]`}}
	ch := planApply(t, psHarness(t, fake), "transportRule.state", engine.Inputs{"rule": "block exe"})
	if ch.Before != "enabled" || ch.After != "disabled" || ch.Target != "Block exe" {
		t.Fatalf("change %+v", ch)
	}
	if w := fake.writes(); len(w) != 1 || w[0].Cmdlet != "Disable-TransportRule" || w[0].Params["Confirm"] != false {
		t.Fatalf("writes %+v", w)
	}
}

func TestSendOnBehalfFallsBackToPowerShell(t *testing.T) {
	fake := &fakePS{answers: map[string]string{"Get-Mailbox": `[{"GrantSendOnBehalfTo":[]}]`}}
	e := psHarness(t, fake)
	for _, c := range e.Catalog() {
		if c.ID == "mailbox.sendOnBehalf" && (!c.Available || c.Backend != pwsh.BackendExchangePS) {
			t.Fatalf("catalog entry %+v", c)
		}
	}
	planApply(t, e, "mailbox.sendOnBehalf", engine.Inputs{"mailbox": "ann@contoso.com", "delegate": "bob@contoso.com"})
	w := fake.writes()
	if hash, _ := w[0].Params["GrantSendOnBehalfTo"].(map[string]any); len(w) != 1 || hash["Add"] != "bob@contoso.com" {
		t.Fatalf("writes %+v", w)
	}
}

func TestEveryPowerShellCmdletIsAllowListed(t *testing.T) {
	allowed := map[string]bool{}
	for _, c := range PowerShellCmdlets()[pwsh.FamilyExchange] {
		allowed[c] = true
	}
	// Exercise every PowerShell action's plan and apply against a recorder.
	fake := &fakePS{answers: map[string]string{
		"Get-Mailbox":            `[{"RecipientTypeDetails":"UserMailbox","EmailAddresses":["SMTP:ann@contoso.com"],"GrantSendOnBehalfTo":[]}]`,
		"Get-TransportRule":      `[{"Name":"r","State":"Enabled"}]`,
		"Get-CalendarProcessing": `[{"AutomateProcessing":"None"}]`,
	}}
	e := psHarness(t, fake)
	for id, in := range map[string]engine.Inputs{
		"mailbox.fullAccess":          {"mailbox": "ann@contoso.com", "user": "bob@contoso.com"},
		"mailbox.sendAs":              {"mailbox": "ann@contoso.com", "user": "bob@contoso.com"},
		"mailbox.type":                {"mailbox": "ann@contoso.com"},
		"mailbox.forwarding":          {"mailbox": "ann@contoso.com", "forwardTo": "bob@contoso.com"},
		"mailbox.address":             {"mailbox": "ann@contoso.com", "address": "x@contoso.com"},
		"mailbox.calendarProcessing":  {"mailbox": "ann@contoso.com"},
		"distributionList.membership": {"user": "bob@contoso.com", "group": "dl1"},
		"transportRule.state":         {"rule": "r"},
		"mailbox.sendOnBehalf":        {"mailbox": "ann@contoso.com", "delegate": "bob@contoso.com"},
		"mailbox.folderPermission":    {"mailbox": "ann@contoso.com", "user": "bob@contoso.com"},
	} {
		planApply(t, e, id, in)
	}
	for _, c := range fake.calls {
		if !allowed[c.Cmdlet] {
			t.Errorf("%s is called but not allow-listed", c.Cmdlet)
		}
	}
}
