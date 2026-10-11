package exoapi

import (
	"context"
	"errors"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// isolate gives a test its own BaseURL, client cache and gone marks, and
// records the retry waits instead of sleeping.
func isolate(t *testing.T, base string) *[]time.Duration {
	t.Helper()
	oldBase, oldSleep := BaseURL, sleep
	clientsMu.Lock()
	oldClients := clients
	clients = map[clientKey]*cachedClient{}
	clientsMu.Unlock()
	goneMu.Lock()
	oldGone := gone
	gone = map[*graphapi.Client]time.Time{}
	goneMu.Unlock()
	var waits []time.Duration
	BaseURL = base
	sleep = func(_ context.Context, d time.Duration) error { waits = append(waits, d); return nil }
	t.Cleanup(func() {
		BaseURL, sleep = oldBase, oldSleep
		clientsMu.Lock()
		clients = oldClients
		clientsMu.Unlock()
		goneMu.Lock()
		gone = oldGone
		goneMu.Unlock()
	})
	return &waits
}

func TestFromEnvReusesOneClientPerConnection(t *testing.T) {
	isolate(t, "https://exo.test")
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

func TestCacheEvictsOnlyTheLeastRecentlyUsedClient(t *testing.T) {
	isolate(t, "https://exo.test")
	envs := make([]engine.Env, maxCachedClients)
	first := make([]*Client, maxCachedClients)
	for i := range envs {
		envs[i] = engine.Env{Tokens: &broker{}, TenantID: "t1", Graph: graphapi.New(graphapi.StaticToken("t"))}
		first[i], _ = FromEnv(envs[i])
	}
	time.Sleep(2 * time.Millisecond) // distinct last-use times on coarse clocks
	for i := 1; i < len(envs); i++ {
		_, _ = FromEnv(envs[i]) // all but the first used again
	}
	_, _ = FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1", Graph: graphapi.New(graphapi.StaticToken("t"))})
	for i := 1; i < len(envs); i++ {
		if c, _ := FromEnv(envs[i]); c != first[i] {
			t.Fatalf("client %d was evicted with the least recently used one", i)
		}
	}
	if c, _ := FromEnv(envs[0]); c == first[0] {
		t.Fatal("the least recently used client was kept past the bound")
	}
}

func TestUnavailableAnswersAreRetriedForPost(t *testing.T) {
	var calls atomic.Int32
	var handler atomic.Value
	handler.Store(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch calls.Add(1) {
		case 1:
			w.WriteHeader(503)
		case 2:
			w.Header().Set("Retry-After", "7")
			w.WriteHeader(504)
		default:
			w.Write([]byte(`{"value":[{"Alias":"a"}]}`))
		}
	}))
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		handler.Load().(http.HandlerFunc)(w, r)
	}))
	defer srv.Close()
	waits := isolate(t, srv.URL)
	c, _ := FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1"})
	items, err := c.Invoke(context.Background(), "Mailbox", "Set-Mailbox", map[string]any{"Identity": "a"}, "UPN:a@b")
	if err != nil || len(items) != 1 || calls.Load() != 3 {
		t.Fatalf("items=%d err=%v calls=%d", len(items), err, calls.Load())
	}
	// Backoff with jitter first, then the gateway's own Retry-After.
	if len(*waits) != 2 || (*waits)[0] < time.Second || (*waits)[0] > 1500*time.Millisecond || (*waits)[1] != 7*time.Second {
		t.Errorf("waits = %v", *waits)
	}

	// A gateway that stays down gives up after the bounded retries.
	calls.Store(0)
	handler.Store(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { calls.Add(1); w.WriteHeader(503) }))
	if _, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); err == nil || calls.Load() != maxUnavailableRetries+1 {
		t.Fatalf("err=%v calls=%d", err, calls.Load())
	}
	// Other failures are not retried here.
	calls.Store(0)
	handler.Store(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { calls.Add(1); w.WriteHeader(500) }))
	if _, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); err == nil || calls.Load() != 1 {
		t.Fatalf("500: err=%v calls=%d", err, calls.Load())
	}
}

func TestCancelDuringRetryWaitReturnsTheContextError(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(503) }))
	defer srv.Close()
	isolate(t, srv.URL)
	ctx, cancel := context.WithCancel(context.Background())
	sleep = func(ctx context.Context, _ time.Duration) error { cancel(); return ctx.Err() }
	c, _ := FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1"})
	if _, err := c.Invoke(ctx, "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); !errors.Is(err, context.Canceled) {
		t.Fatalf("err = %v, want context.Canceled", err)
	}
}

func TestChangedEndpointSwitchesConnectionToPowerShell(t *testing.T) {
	var cmdletError atomic.Bool
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if cmdletError.Load() {
			w.WriteHeader(400)
			w.Write([]byte(`{"error":{"code":"BadRequest","message":"|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object 'x' couldn't be found."}}`))
			return
		}
		w.WriteHeader(404)
		w.Write([]byte(`Not Found`))
	}))
	defer srv.Close()
	isolate(t, srv.URL)

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
	cmdletError.Store(true)
	c, _ := FromEnv(env)
	if _, err := c.Invoke(env.Ctx, "Mailbox", "Get-Mailbox", map[string]any{"Identity": "x"}, "UPN:x@b"); err == nil {
		t.Fatal("cmdlet error swallowed")
	}
	if r := p.Status(s); r != nil {
		t.Fatalf("a cmdlet error switched the backend: %+v", r)
	}

	// The endpoint itself is gone: the error says to run again, and the
	// resolver's next status check moves this connection to PowerShell.
	cmdletError.Store(false)
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
