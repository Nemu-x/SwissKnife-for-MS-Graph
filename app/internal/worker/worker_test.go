package worker

import (
	"context"
	"encoding/json"
	"errors"
	"net"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/pwsh"
)

type fakeRunner struct {
	calls []string
	token string
}

func (f *fakeRunner) Invoke(env engine.Env, family, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	f.calls = append(f.calls, family+":"+cmdlet)
	f.token, _ = env.Tokens.TokenFor(env.Ctx, auth.ResourceExchange)
	if cmdlet == "Get-Fails" {
		return nil, &pwsh.Error{Message: "boom", ErrorID: "X"}
	}
	b, _ := json.Marshal(map[string]any{"cmdlet": cmdlet, "identity": params["Identity"]})
	return []json.RawMessage{b}, nil
}

func startWorker(t *testing.T) (*Server, string, *fakeRunner) {
	t.Helper()
	dir := t.TempDir()
	id, err := LoadOrCreateIdentity(dir, "worker")
	if err != nil {
		t.Fatal(err)
	}
	r := &fakeRunner{}
	s := &Server{Identity: id, Dir: dir, Name: "SRV01", Version: "test", Runner: r,
		Families: map[string]bool{"exo": true}, Allow: map[string][]string{"exo": {"Get-Mailbox", "Get-Fails"}, "ipps": {"Get-ComplianceSearch"}},
		Audit: auditlog.New(dir)}
	l, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	t.Cleanup(cancel)
	go func() { _ = s.ServeListener(ctx, l) }()
	return s, l.Addr().String(), r
}

func TestPairThenInvokeOverPinnedMTLS(t *testing.T) {
	s, addr, runner := startWorker(t)
	client, err := LoadOrCreateIdentity(t.TempDir(), "client")
	if err != nil {
		t.Fatal(err)
	}
	ctx := context.Background()

	// Before pairing the handshake itself refuses an unknown certificate.
	if _, err := NewRemote(addr, s.Identity.Fingerprint, client).Health(ctx); err == nil {
		t.Fatal("an unpaired client must be refused")
	}

	code := s.StartPairing()
	if _, _, err := Pair(ctx, addr, "WRONG-CODE-00", "mac", client); err == nil {
		t.Fatal("a wrong code must not pair")
	}
	fp, name, err := Pair(ctx, addr, strings.ToLower(code), "Ann's Mac", client)
	if err != nil {
		t.Fatal(err)
	}
	if fp != s.Identity.Fingerprint || name != "SRV01" {
		t.Fatalf("pinned %s name %s", fp, name)
	}
	if _, _, err := Pair(ctx, addr, code, "again", client); err == nil {
		t.Fatal("a code works once")
	}

	r := NewRemote(addr, fp, client)
	h, err := r.Health(ctx)
	if err != nil || len(h.Families) != 1 || h.Families[0] != "exo" {
		t.Fatalf("health %+v %v", h, err)
	}
	req := InvokeRequest{Family: "exo", Cmdlet: "Get-Mailbox", Params: map[string]any{"Identity": "ann@contoso.com"},
		TenantID: "t1", AppOnly: true, Tokens: map[string]string{auth.ResourceExchange: "exo-token", auth.ResourceGraph: "g"}}
	out, err := r.Invoke(ctx, req)
	if err != nil || len(out) != 1 || !strings.Contains(string(out[0]), "ann@contoso.com") || runner.token != "exo-token" {
		t.Fatalf("invoke %s %v token %q", out, err, runner.token)
	}

	// PowerShell errors keep their type; cmdlets outside the worker's
	// allow-list or families never reach PowerShell.
	req.Cmdlet = "Get-Fails"
	var pe *pwsh.Error
	if _, err := r.Invoke(ctx, req); !errors.As(err, &pe) || pe.ErrorID != "X" {
		t.Fatalf("error %v", err)
	}
	for _, c := range []InvokeRequest{{Family: "exo", Cmdlet: "Remove-Mailbox"}, {Family: "ipps", Cmdlet: "Get-ComplianceSearch"}} {
		if _, err := r.Invoke(ctx, c); err == nil || !strings.Contains(err.Error(), "does not run") {
			t.Fatalf("%s must be refused: %v", c.Cmdlet, err)
		}
	}
	if len(runner.calls) != 2 {
		t.Fatalf("calls %v", runner.calls)
	}

	// A pinned worker fingerprint that does not match is refused too.
	if _, err := NewRemote(addr, strings.Repeat("0", 64), client).Health(ctx); err == nil {
		t.Fatal("a different worker certificate must be refused")
	}

	// Revoking unpairs; tokens never reach the disk.
	if err := s.Revoke(client.Fingerprint[:12]); err != nil {
		t.Fatal(err)
	}
	if _, err := r.Health(ctx); err == nil {
		t.Fatal("a revoked client must be refused")
	}
	_ = filepath.WalkDir(s.Dir, func(p string, d os.DirEntry, _ error) error {
		if !d.IsDir() {
			if b, _ := os.ReadFile(p); strings.Contains(string(b), "exo-token") {
				t.Fatalf("token written to %s", p)
			}
		}
		return nil
	})
}
