package exoapi

import (
	"context"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

func noSleep(t *testing.T) *[]time.Duration {
	t.Helper()
	var waits []time.Duration
	old := sleep
	sleep = func(_ context.Context, d time.Duration) error { waits = append(waits, d); return nil }
	t.Cleanup(func() { sleep = old })
	return &waits
}

func TestFromEnvReusesOneClientPerConnection(t *testing.T) {
	BaseURL = "https://exo.test"
	b := &broker{}
	conn := graphapi.New(graphapi.StaticToken("t"))
	env := engine.Env{Tokens: b, TenantID: "t1", Graph: conn}
	c1, _ := FromEnv(env)
	c2, _ := FromEnv(env)
	if c1 != c2 {
		t.Fatal("a new client was built for the same connection")
	}
	env.Graph = graphapi.New(graphapi.StaticToken("t")) // reconnect
	if c3, _ := FromEnv(env); c3 == c1 {
		t.Fatal("a new connection reused the old client")
	}
	if c4, _ := FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1", Graph: conn}); c4 == c1 {
		t.Fatal("another token broker reused the client")
	}
}

func TestUnavailableAnswersAreRetriedForPost(t *testing.T) {
	waits := noSleep(t)
	calls := 0
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		calls++
		if calls <= 2 {
			w.WriteHeader([]int{503, 504}[calls-1])
			return
		}
		w.Write([]byte(`{"value":[{"Alias":"a"}]}`))
	}))
	defer srv.Close()
	BaseURL = srv.URL
	c, _ := FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1"})
	items, err := c.Invoke(context.Background(), "Mailbox", "Set-Mailbox", map[string]any{"Identity": "a"}, "UPN:a@b")
	if err != nil || len(items) != 1 || calls != 3 {
		t.Fatalf("items=%d err=%v calls=%d", len(items), err, calls)
	}
	if len(*waits) != 2 || (*waits)[1] <= (*waits)[0] {
		t.Errorf("backoff = %v", *waits)
	}

	// A gateway that stays down gives up after the bounded retries.
	calls = 0
	srv.Config.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { calls++; w.WriteHeader(503) })
	if _, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); err == nil || calls != maxUnavailableRetries+1 {
		t.Fatalf("err=%v calls=%d", err, calls)
	}
	// Other failures are not retried here.
	calls = 0
	srv.Config.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { calls++; w.WriteHeader(500) })
	if _, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); err == nil || calls != 1 {
		t.Fatalf("500: err=%v calls=%d", err, calls)
	}
}

func TestChangedEndpointSwitchesConnectionToPowerShell(t *testing.T) {
	noSleep(t)
	cmdletError := false
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if cmdletError {
			w.WriteHeader(400)
			w.Write([]byte(`{"error":{"code":"BadRequest","message":"|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object 'x' couldn't be found."}}`))
			return
		}
		w.WriteHeader(404)
		w.Write([]byte(`Not Found`))
	}))
	defer srv.Close()
	BaseURL = srv.URL

	s := session.New(auditlog.New(t.TempDir()))
	conn := graphapi.New(graphapi.StaticToken("t"))
	s.SetClient(conn, "a")
	s.SetTokens(&broker{})
	s.SetIdentity("t1", true)
	p := NewProvider()
	p.probe = func(context.Context, engine.Env) error { return nil }
	if r := p.Status(s); r != nil {
		t.Fatalf("available before: %+v", r)
	}
	env := engine.Env{Ctx: context.Background(), Graph: conn, Tokens: s.Tokens(), TenantID: "t1"}

	// A mailbox that does not exist is the operator's input, not the API.
	cmdletError = true
	c, _ := FromEnv(env)
	if _, err := c.Invoke(env.Ctx, "Mailbox", "Get-Mailbox", map[string]any{"Identity": "x"}, "UPN:x@b"); err == nil {
		t.Fatal("cmdlet error swallowed")
	}
	if r := p.Status(s); r != nil {
		t.Fatalf("a cmdlet error switched the backend: %+v", r)
	}

	// The endpoint itself is gone: the error says to run again, and the
	// resolver's next status check moves this connection to PowerShell.
	cmdletError = false
	_, err := c.Invoke(env.Ctx, "MailboxFolderPermission", "Get-MailboxFolderPermission", nil, "UPN:a@b")
	if err == nil || !strings.Contains(err.Error(), "run the action again") {
		t.Fatalf("err = %v", err)
	}
	if r := p.Status(s); r == nil || r.Key != "exoApiChanged" {
		t.Fatalf("status after a gone endpoint = %+v", r)
	}
	// A new connection probes afresh.
	s.SetClient(graphapi.New(graphapi.StaticToken("t")), "b")
	if r := p.Status(s); r != nil {
		t.Fatalf("new connection: %+v", r)
	}
}

func TestEndpointGoneClassification(t *testing.T) {
	cases := []struct {
		err  error
		gone bool
	}{
		{&graphapi.GraphError{StatusCode: 501}, true},
		{&graphapi.GraphError{StatusCode: 404, Message: "Not Found"}, true},
		{&graphapi.GraphError{StatusCode: 400, Message: "Unknown endpoint"}, true},
		{&graphapi.GraphError{StatusCode: 404, Message: "|Microsoft.Exchange.Data.ManagementObjectNotFoundException|nope"}, false},
		{&graphapi.GraphError{StatusCode: 400, Message: "Cannot bind parameter: ParameterBindingException"}, false},
		{&graphapi.GraphError{StatusCode: 403}, false},
		{&graphapi.GraphError{StatusCode: 503}, false},
	}
	for _, tc := range cases {
		if got := endpointGone(tc.err); got != tc.gone {
			t.Errorf("%v: gone=%v, want %v", tc.err, got, tc.gone)
		}
	}
}
