package exoapi

import (
	"context"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"testing"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

type broker struct{ resources []string }

func (b *broker) TokenFor(_ context.Context, resource string) (string, error) {
	b.resources = append(b.resources, resource)
	return "exo-token", nil
}

func TestInvokeFollowsNextLinkWithSameBodyAndAnchor(t *testing.T) {
	var bodies []string
	var srvURL string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		bodies = append(bodies, string(b))
		if r.Method != "POST" || r.Header.Get("X-AnchorMailbox") != "UPN:ann@contoso.com" || r.Header.Get("Authorization") != "Bearer exo-token" {
			t.Errorf("request %s anchor=%q auth=%q", r.Method, r.Header.Get("X-AnchorMailbox"), r.Header.Get("Authorization"))
		}
		if r.URL.Path == "/adminapi/v2.0/t1/Mailbox" && r.URL.Query().Get("page") == "" {
			w.Write([]byte(`{"value":[{"Alias":"a"}],"@odata.nextLink":"` + srvURL + `/adminapi/v2.0/t1/Mailbox?page=2"}`))
			return
		}
		w.Write([]byte(`{"value":[{"Alias":"b"}]}`))
	}))
	defer srv.Close()
	srvURL = srv.URL
	BaseURL = srv.URL
	b := &broker{}
	c, err := FromEnv(engine.Env{Tokens: b, TenantID: "t1"})
	if err != nil {
		t.Fatal(err)
	}
	items, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", map[string]any{"ResultSize": 1}, MailboxAnchor("ann@contoso.com"))
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 2 || len(bodies) != 2 || bodies[0] != bodies[1] {
		t.Fatalf("items %d bodies %v", len(items), bodies)
	}
	var in struct {
		CmdletInput struct {
			CmdletName string
			Parameters map[string]any
		}
	}
	_ = json.Unmarshal([]byte(bodies[0]), &in)
	if in.CmdletInput.CmdletName != "Get-Mailbox" || in.CmdletInput.Parameters["ResultSize"] != float64(1) {
		t.Fatalf("body %s", bodies[0])
	}
	if len(b.resources) == 0 || b.resources[0] != "https://outlook.office365.com" {
		t.Fatalf("token resource %v", b.resources)
	}
}

func TestDecodeShapes(t *testing.T) {
	for raw, want := range map[string]int{``: 0, `null`: 0, `[{"a":1},{"a":2}]`: 2, `{"Name":"x"}`: 1, `{"value":[]}`: 0} {
		if got, _, err := decode(json.RawMessage(raw)); err != nil || len(got) != want {
			t.Errorf("decode(%q) = %d items, want %d", raw, len(got), want)
		}
	}
}

func TestClassify(t *testing.T) {
	cases := map[int]string{403: "exoPermission", 401: "exoPermission", 404: "exoApiNotEnabled", 400: "exoApiNotEnabled", 500: "exoUnreachable"}
	for code, key := range cases {
		if r := classify(&graphapi.GraphError{StatusCode: code}); r == nil || r.Key != key {
			t.Errorf("%d → %+v, want %s", code, r, key)
		}
	}
	if classify(nil) != nil {
		t.Error("success must be available")
	}
	if r := classify(errors.New("dial tcp: timeout")); r.Key != "exoUnreachable" {
		t.Errorf("network error → %+v", r)
	}
}

func TestProviderProbesOncePerConnection(t *testing.T) {
	s := session.New(auditlog.New(t.TempDir()))
	p := NewProvider()
	calls := 0
	p.probe = func(context.Context, engine.Env) error { calls++; return nil }
	if r := p.Status(s); r == nil || r.Key != "notConnected" {
		t.Fatalf("disconnected: %+v", r)
	}
	s.SetClient(graphapi.New(graphapi.StaticToken("t")), "a")
	s.SetTokens(&broker{})
	s.SetIdentity("t1", true)
	p.Status(s)
	p.Status(s)
	if calls != 1 {
		t.Fatalf("probe ran %d times, want 1", calls)
	}
	s.SetClient(graphapi.New(graphapi.StaticToken("t")), "b") // reconnect
	p.Status(s)
	if calls != 2 {
		t.Fatalf("reconnect must re-probe: %d", calls)
	}
}

func TestTransientProbeFailureIsRetried(t *testing.T) {
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t")), "a")
	s.SetTokens(&broker{})
	s.SetIdentity("t1", true)
	p := NewProvider()
	calls := 0
	p.probe = func(context.Context, engine.Env) error { calls++; return transient{errors.New("graph down")} }
	now := time.Now()
	p.now = func() time.Time { return now }
	if r := p.Status(s); r == nil || r.Key != "exoUnreachable" {
		t.Fatalf("status %+v", r)
	}
	p.Status(s)
	if calls != 1 {
		t.Fatalf("within the retry window: %d probes", calls)
	}
	now = now.Add(2 * time.Minute)
	p.Status(s)
	if calls != 2 {
		t.Fatalf("after the window a transient failure must be re-probed: %d", calls)
	}
}

func TestNextLinkToAnotherHostIsRefused(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"value":[],"@odata.nextLink":"https://evil.example/steal"}`))
	}))
	defer srv.Close()
	BaseURL = srv.URL
	c, _ := FromEnv(engine.Env{Tokens: &broker{}, TenantID: "t1"})
	if _, err := c.Invoke(context.Background(), "Mailbox", "Get-Mailbox", nil, "UPN:a@b"); err == nil {
		t.Fatal("a continuation to another host must be refused")
	}
}
