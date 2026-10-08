package services

import (
	"encoding/json"
	"net/http"
	"strings"
	"testing"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// recordingPS stands in for Exchange PowerShell: it answers Get-Mailbox and
// records every cmdlet the playbook runs.
type recordingPS struct{ calls []string }

func (r *recordingPS) Invoke(_ engine.Env, _, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	b, _ := json.Marshal(params)
	r.calls = append(r.calls, cmdlet+" "+string(b))
	switch cmdlet {
	case "Get-Mailbox":
		return []json.RawMessage{json.RawMessage(`{"RecipientTypeDetails":"UserMailbox"}`)}, nil
	}
	return nil, nil
}

type readyProvider struct{ b engine.Backend }

func (p readyProvider) Backend() engine.Backend                { return p.b }
func (p readyProvider) Status(*session.Session) *engine.Reason { return nil }

// The leaver's mailbox becomes shared and opens to the manager BEFORE the
// licenses go — the order that keeps the mail — and Graph's lagging mailbox
// type does not fail the run once the conversion succeeded.
func TestOffboardConvertsToSharedAndGrantsAccessBeforeLicenses(t *testing.T) {
	var graph []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		graph = append(graph, r.Method+" "+r.URL.Path)
		switch {
		case strings.HasSuffix(r.URL.Path, "/licenseDetails"):
			w.Write([]byte(`{"value":[{"skuId":"s1"}]}`))
		case strings.HasSuffix(r.URL.Path, "/mailboxSettings"):
			w.Write([]byte(`{"userPurpose":"user"}`)) // not replicated yet
		case strings.HasPrefix(r.URL.Path, "/users/"):
			w.Write([]byte(`{"id":"u1","userPrincipalName":"` + strings.TrimPrefix(r.URL.Path, "/users/") + `","mail":"` + strings.TrimPrefix(r.URL.Path, "/users/") + `"}`))
		default:
			w.Write([]byte(`{}`))
		}
	})
	ps := &recordingPS{}
	e := engine.New(sess, engine.GraphProvider{}, readyProvider{pwsh.BackendExchangePS})
	e.PS = ps
	e.Register(actions.Builtin()...)
	engines.Store(sess, e)
	t.Cleanup(func() { engines.Delete(sess) })

	res, err := NewPlaybookService(sess).Offboard(OffboardRequest{
		Upn: "leaver@contoso.com", Confirm: "leaver@contoso.com",
		ConvertToShared: true, FullAccessTo: "boss@contoso.com", RemoveAllLicenses: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if !res.OK {
		t.Fatalf("run failed: %+v", res.Steps)
	}
	var names []string
	for _, s := range res.Steps {
		names = append(names, s.Name)
	}
	want := []string{"Convert to shared mailbox", "Grant mailbox access", "Check mailbox type", "Remove licenses"}
	if strings.Join(names, ",") != strings.Join(want, ",") {
		t.Fatalf("steps %v, want %v", names, want)
	}
	joined := strings.Join(ps.calls, "\n")
	if !strings.Contains(joined, `Set-Mailbox {"Identity":"leaver@contoso.com","Type":"Shared"}`) ||
		!strings.Contains(joined, "Add-MailboxPermission") {
		t.Fatalf("powershell calls:\n%s", joined)
	}
	for _, c := range graph {
		if strings.HasSuffix(c, "/mailboxSettings") {
			t.Errorf("the mailbox-type pre-flight must be skipped after a conversion: %s", c)
		}
	}
}
