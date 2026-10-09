package services

import (
	"context"
	"encoding/json"
	"net"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
	"swissknife-app/internal/worker"
)

type echoRunner struct {
	token, graphToken, org string
}

func (r *echoRunner) Invoke(env engine.Env, family, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	r.token, _ = env.Tokens.TokenFor(env.Ctx, auth.ResourceExchange)
	r.graphToken, _ = env.Tokens.TokenFor(env.Ctx, auth.ResourceGraph)
	r.org = env.ExchangeOrg
	b, _ := json.Marshal(map[string]any{"ran": cmdlet, "on": "worker", "tenant": env.TenantID})
	return []json.RawMessage{b}, nil
}

type tokenBroker struct{}

func (tokenBroker) TokenFor(_ context.Context, resource string) (string, error) { return "tok-" + resource, nil }

// A machine without PowerShell pairs with a worker; an Exchange cmdlet call
// then runs there, with a fresh token, and is audited on both sides.
func TestCmdletsGoToThePairedWorkerWhenThisMachineCannot(t *testing.T) {
	wdir := t.TempDir()
	wid, _ := worker.LoadOrCreateIdentity(wdir, "worker")
	runner := &echoRunner{}
	srv := &worker.Server{Identity: wid, Dir: wdir, Name: "SRV01", NewRunner: func() engine.PSRunner { return runner },
		Families: map[string]bool{"exo": true}, Allow: map[string][]string{"exo": {"Get-Mailbox"}}, Audit: auditlog.New(wdir)}
	l, _ := net.Listen("tcp", "127.0.0.1:0")
	ctx, cancel := context.WithCancel(context.Background())
	t.Cleanup(cancel)
	go func() { _ = srv.ServeListener(ctx, l) }()
	code := srv.StartPairing()

	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(t.TempDir())
	sess.SetClient(graphapi.New(graphapi.StaticToken("t")), "test")
	sess.SetTokens(tokenBroker{})
	sess.SetIdentity("tenant-1", true)
	ws := NewWorkerService(sess)
	w, err := ws.Pair(l.Addr().String(), code)
	if err != nil || w.Name != "SRV01" || len(w.Families) != 1 {
		t.Fatalf("pair: %+v %v", w, err)
	}

	d := &psDispatch{s: sess, local: pwsh.NewPool(pwsh.NewDetector(), nil), can: func(string) bool { return false }}
	graph := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(`{"value":[{"verifiedDomains":[{"name":"contoso.onmicrosoft.com","isInitial":true}]}]}`))
	}))
	t.Cleanup(graph.Close)
	gc := graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(graph.URL))
	env := engine.Env{Ctx: context.Background(), Graph: gc, Tokens: tokenBroker{}, TenantID: "tenant-1", AppOnly: true}
	out, err := d.Invoke(env, pwsh.FamilyExchange, "Get-Mailbox", map[string]any{"Identity": "ann"})
	if err != nil || len(out) != 1 || !strings.Contains(string(out[0]), `"worker"`) || runner.token != "tok-"+auth.ResourceExchange {
		t.Fatalf("remote call: %s %v token %q", out, err, runner.token)
	}
	// Exchange signs in with the org found here: the worker gets no Graph token.
	if runner.org != "contoso.onmicrosoft.com" || runner.graphToken != "" {
		t.Fatalf("org %q graph token %q", runner.org, runner.graphToken)
	}

	// The provider counts the worker's families as available here.
	p := workerAware{inner: unavailablePS{}, family: pwsh.FamilyExchange}
	if r := p.Status(sess); r != nil {
		t.Fatalf("worker-backed provider: %+v", r)
	}
	if r := (workerAware{inner: unavailablePS{}, family: pwsh.FamilyTeams}).Status(sess); r == nil {
		t.Fatal("a family the worker does not run stays unavailable")
	}
	if err := ws.Unpair(); err != nil || ws.Status() != nil {
		t.Fatalf("unpair: %v", err)
	}
}

type unavailablePS struct{}

func (unavailablePS) Backend() engine.Backend { return pwsh.BackendExchangePS }
func (unavailablePS) Status(*session.Session) *engine.Reason {
	return &engine.Reason{Key: "pwshMissing"}
}

// The worker card's status: online with a last-seen time after a check;
// offline (with why) once it stops answering, keeping what it ran.
func TestWorkerCheckReportsOnlineAndOffline(t *testing.T) {
	wdir := t.TempDir()
	wid, _ := worker.LoadOrCreateIdentity(wdir, "worker")
	srv := &worker.Server{Identity: wid, Dir: wdir, Name: "SRV01", Version: "2.0.0", NewRunner: func() engine.PSRunner { return &echoRunner{} },
		Families: map[string]bool{"exo": true, "teams": true}, Available: func() map[string]bool { return map[string]bool{"exo": true} },
		Allow: map[string][]string{"exo": {"Get-Mailbox"}}, Audit: auditlog.New(wdir)}
	l, _ := net.Listen("tcp", "127.0.0.1:0")
	ctx, cancel := context.WithCancel(context.Background())
	go func() { _ = srv.ServeListener(ctx, l) }()
	code := srv.StartPairing()

	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(t.TempDir())
	ws := NewWorkerService(sess)
	if _, err := ws.Pair(l.Addr().String(), code); err != nil {
		t.Fatal(err)
	}
	w, err := ws.Check()
	if err != nil || !w.Online || w.LastSeen.IsZero() || w.Version != "2.0.0" ||
		strings.Join(w.Families, ",") != "exo" || strings.Join(w.Unavailable, ",") != "teams" {
		t.Fatalf("online: %+v %v", w, err)
	}
	if ws.Status().Online {
		t.Fatal("online is known only from a check in this session")
	}
	cancel()
	_ = l.Close()
	workerMu.Lock()
	for k, r := range remotes { // drop the kept-alive connection
		r.Close()
		delete(remotes, k)
	}
	workerMu.Unlock()
	w, err = ws.Check()
	if err != nil || w.Online || w.LastError == "" || strings.Join(w.Families, ",") != "exo" || w.LastSeen.IsZero() {
		t.Fatalf("offline: %+v %v", w, err)
	}
}
