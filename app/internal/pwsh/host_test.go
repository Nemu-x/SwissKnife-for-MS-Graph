package pwsh

import (
	"bufio"
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// The test binary doubles as a fake pwsh: with SKG_FAKE_PWSH set it speaks
// the host protocol on stdin/stdout instead of running tests. Module noise
// on stdout (banners, warnings) is simulated too.
func TestMain(m *testing.M) {
	if os.Getenv("SKG_FAKE_PWSH") == "1" {
		fakeHost()
		return
	}
	os.Exit(m.Run())
}

func fakeHost() {
	allow := map[string]bool{}
	connects, authFails := 0, 0
	in := bufio.NewScanner(os.Stdin)
	reply := func(v any) {
		b, _ := json.Marshal(v)
		fmt.Println("WARNING: a module banner that is not a reply")
		fmt.Print(marker + string(b) + "\n")
	}
	for in.Scan() {
		var req map[string]any
		_ = json.Unmarshal(in.Bytes(), &req)
		id := req["id"]
		switch req["op"] {
		case "init":
			for _, c := range req["allow"].([]any) {
				allow[c.(string)] = true
			}
			reply(map[string]any{"id": id, "ok": true, "data": []any{}})
		case "connect":
			connects++
			if req["token"] == "" {
				reply(map[string]any{"id": id, "ok": false, "error": map[string]any{"message": "no token"}})
				continue
			}
			reply(map[string]any{"id": id, "ok": true, "data": []any{}})
		case "invoke":
			cmdlet := req["cmdlet"].(string)
			switch {
			case !allow[cmdlet]:
				reply(map[string]any{"id": id, "ok": false, "error": map[string]any{"message": "cmdlet not allowed: " + cmdlet}})
			case cmdlet == "Start-Sleep":
				time.Sleep(time.Minute)
			case cmdlet == "Stop-Process":
				os.Exit(3)
			case cmdlet == "Get-Expiring" && authFails == 0:
				// The first call after the first sign-in fails like an expired token.
				authFails++
				reply(map[string]any{"id": id, "ok": false, "error": map[string]any{"message": "The remote server returned 401 Unauthorized"}})
			case cmdlet == "Get-Connects":
				reply(map[string]any{"id": id, "ok": true, "data": []any{connects}})
			default:
				reply(map[string]any{"id": id, "ok": true, "data": []any{map[string]any{"cmdlet": cmdlet, "params": req["params"]}}})
			}
		}
	}
}

func startFake(t *testing.T, allow ...string) *Host {
	t.Helper()
	t.Setenv("SKG_FAKE_PWSH", "1")
	exe, err := os.Executable()
	if err != nil {
		t.Fatal(err)
	}
	h, err := Start(exe, allow)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(h.Close)
	return h
}

func TestInvokeRoundTripIgnoresModuleNoise(t *testing.T) {
	h := startFake(t, "Get-Mailbox")
	if err := h.Connect(context.Background(), "exo", map[string]any{"token": "t", "organization": "contoso.onmicrosoft.com"}); err != nil {
		t.Fatal(err)
	}
	out, err := h.Invoke(context.Background(), "Get-Mailbox", map[string]any{"Identity": "ann; Remove-Mailbox x"})
	if err != nil {
		t.Fatal(err)
	}
	var got struct {
		Cmdlet string
		Params map[string]string
	}
	if len(out) != 1 || json.Unmarshal(out[0], &got) != nil || got.Params["Identity"] != "ann; Remove-Mailbox x" {
		t.Fatalf("out %s", out)
	}
}

func TestRefusedCmdletIsAnError(t *testing.T) {
	h := startFake(t, "Get-Mailbox")
	_, err := h.Invoke(context.Background(), "Remove-Mailbox", nil)
	var pe *Error
	if !errors.As(err, &pe) || !strings.Contains(pe.Message, "not allowed") {
		t.Fatalf("err = %v", err)
	}
}

func TestCancelKillsTheHost(t *testing.T) {
	h := startFake(t, "Start-Sleep")
	ctx, cancel := context.WithTimeout(context.Background(), 200*time.Millisecond)
	defer cancel()
	start := time.Now()
	if _, err := h.Invoke(ctx, "Start-Sleep", nil); !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("err = %v", err)
	}
	if time.Since(start) > 5*time.Second {
		t.Fatal("cancel did not return promptly")
	}
	deadline := time.Now().Add(5 * time.Second)
	for h.Alive() && time.Now().Before(deadline) {
		time.Sleep(20 * time.Millisecond)
	}
	if h.Alive() {
		t.Fatal("the host must be killed on cancel")
	}
}

func TestCrashIsReportedAsExit(t *testing.T) {
	h := startFake(t, "Stop-Process")
	if _, err := h.Invoke(context.Background(), "Stop-Process", nil); !errors.Is(err, ErrHostExited) {
		t.Fatalf("err = %v", err)
	}
}

func TestVersionCompare(t *testing.T) {
	cases := []struct {
		v    string
		min  []int
		want bool
	}{
		{"7.4.6", []int{7, 2}, true},
		{"7.2.0", []int{7, 2}, true},
		{"7.1.9", []int{7, 2}, false},
		{"3.9.2", []int{3, 1}, true},
		{"3.0.0", []int{3, 1}, false},
		{"7.5.0-preview.3", []int{7, 2}, true},
		{"", []int{7, 2}, false},
	}
	for _, c := range cases {
		if got := atLeast(c.v, c.min); got != c.want {
			t.Errorf("atLeast(%q, %v) = %v", c.v, c.min, got)
		}
	}
}

func TestDetectorParsesAndCaches(t *testing.T) {
	runs := 0
	d := &Detector{
		find: func() string { return "/usr/bin/pwsh" },
		run: func(context.Context, string, ...string) ([]byte, error) {
			runs++
			return []byte(`{"version":"7.4.6","modules":{"ExchangeOnlineManagement":"3.9.2"}}`), nil
		},
	}
	env := d.Get(context.Background())
	if !env.PwshOK() || !env.Supports(ModuleExchange) || env.Supports(ModuleTeams) {
		t.Fatalf("env %+v", env)
	}
	d.Get(context.Background())
	if runs != 1 {
		t.Fatalf("detection ran %d times", runs)
	}
	d.Invalidate()
	d.Get(context.Background())
	if runs != 2 {
		t.Fatalf("invalidate must re-detect: %d", runs)
	}
	if err := d.Install(context.Background(), "Evil; rm -rf /"); err == nil {
		t.Fatal("unknown modules must be refused")
	}
}

func fakePool(t *testing.T, allow ...string) *Pool {
	t.Helper()
	t.Setenv("SKG_FAKE_PWSH", "1")
	exe, _ := os.Executable()
	det := &Detector{find: func() string { return exe }, run: func(context.Context, string, ...string) ([]byte, error) {
		return []byte(`{"version":"7.4.0","modules":{}}`), nil
	}}
	p := NewPool(det, map[string][]string{FamilyExchange: allow})
	t.Cleanup(p.Close)
	return p
}

type staticTokens struct{}

func (staticTokens) TokenFor(context.Context, string) (string, error) { return "tok", nil }

func poolEnv(t *testing.T) engine.Env {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"value":[{"verifiedDomains":[{"name":"contoso.onmicrosoft.com","isInitial":true}]}]}`))
	}))
	t.Cleanup(srv.Close)
	return engine.Env{Ctx: context.Background(), Tokens: staticTokens{}, AppOnly: true, TenantID: "t",
		Graph: graphapi.New(graphapi.StaticToken("g"), graphapi.WithBaseURL(srv.URL))}
}

func TestPoolRetriesOnceAfterAnAuthFailure(t *testing.T) {
	p := fakePool(t, "Get-Expiring", "Get-Connects")
	env := poolEnv(t)
	if _, err := p.Invoke(env, FamilyExchange, "Get-Expiring", nil); err != nil {
		t.Fatalf("the retry after signing in again must succeed: %v", err)
	}
	out, err := p.Invoke(env, FamilyExchange, "Get-Connects", nil)
	if err != nil || string(out[0]) != "2" {
		t.Fatalf("connects = %s, %v (want a second sign-in)", out, err)
	}
}

func TestPoolDiscardsHostStartedBeforeClose(t *testing.T) {
	p := fakePool(t, "Get-Connects")
	inner := p.start
	p.start = func(exe string, allow, _ []string) (*Host, error) {
		h, err := inner(exe, allow, nil)
		p.gen++ // a Close landed while the host was starting
		return h, err
	}
	if _, err := p.Invoke(poolEnv(t), FamilyExchange, "Get-Connects", nil); err == nil {
		t.Fatal("a host started across a Close must not be used")
	}
	if len(p.hosts) != 0 {
		t.Fatal("it must not be kept either")
	}
}

func TestSignInValidUntilReadsExp(t *testing.T) {
	now := time.Unix(1_000_000, 0)
	payload := base64.RawURLEncoding.EncodeToString([]byte(`{"exp":1003600}`))
	got := signInValidUntil("h."+payload+".s", now)
	if want := time.Unix(1003600, 0).Add(-renewBefore); !got.Equal(want) {
		t.Fatalf("got %v want %v", got, want)
	}
	if got := signInValidUntil("opaque", now); !got.Equal(now.Add(fallbackTTL)) {
		t.Fatalf("fallback %v", got)
	}
}

func TestCloseForLeavesANewerConnection(t *testing.T) {
	p := fakePool(t, "Get-Connects")
	oldEnv, newEnv := poolEnv(t), poolEnv(t)
	if _, err := p.Invoke(newEnv, FamilyExchange, "Get-Connects", nil); err != nil {
		t.Fatal(err)
	}
	p.CloseFor(oldEnv.Graph) // a late cleanup of the previous connection
	if len(p.hosts) != 1 {
		t.Fatal("the newer connection's host must survive")
	}
	if _, err := p.Invoke(oldEnv, FamilyExchange, "Get-Connects", nil); err == nil {
		t.Fatal("a closed connection must not get a host again")
	}
	if e := p.hosts[FamilyExchange]; e == nil || e.conn != newEnv.Graph || !e.host.Alive() {
		t.Fatal("the rejected call must leave the newer connection's host alone")
	}
}
