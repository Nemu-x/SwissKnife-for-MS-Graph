package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

// fake wires the command layer to buffers, an in-memory profile list and an
// httptest Graph. The production connect path (profiles.json + OS keychain +
// azidentity) is not exercised here: it needs a keychain and a tenant.
type fake struct {
	e        *env
	out, err bytes.Buffer
	picked   []string // profile ids connect was asked for
	readOnly bool
	srv      *httptest.Server
	handler  http.HandlerFunc // set per test; nil answers {}
	calls    []string         // "METHOD /path?query" of every Graph request
}

func newFake(t *testing.T, profiles ...secrets.Profile) *fake {
	t.Helper()
	f := &fake{}
	f.srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		f.calls = append(f.calls, r.Method+" "+r.URL.RequestURI())
		if f.handler != nil {
			f.handler(w, r)
			return
		}
		w.Write([]byte(`{}`))
	}))
	t.Cleanup(f.srv.Close)
	f.e = &env{
		stdout:  &f.out,
		stderr:  &f.err,
		version: "1.2.3-test",
		ctx:     context.Background(),
		profiles: func() ([]secrets.Profile, error) {
			return profiles, nil
		},
		connect: func(ctx context.Context, p secrets.Profile, readOnly bool, _ io.Writer) (*session.Session, error) {
			f.picked = append(f.picked, p.ID)
			f.readOnly = readOnly
			sess := session.New(auditlog.New(t.TempDir()))
			sess.SetAppContext(ctx)
			sess.SetReadOnly(readOnly)
			sess.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(f.srv.URL)), p.Name)
			return sess, nil
		},
	}
	return f
}

var one = secrets.Profile{ID: "p1", Name: "Prod", TenantID: "t1", ClientID: "c1", AuthMode: "client_secret"}
var two = secrets.Profile{ID: "p2", Name: "Lab", TenantID: "t2", ClientID: "c2", AuthMode: "device_code"}

func TestVersion(t *testing.T) {
	f := newFake(t)
	if code := run(f.e, []string{"version"}); code != exitOK {
		t.Fatalf("exit %d, stderr %q", code, f.err.String())
	}
	if got := f.out.String(); got != "SwissKnifeGraph 1.2.3-test\n" {
		t.Fatalf("stdout %q", got)
	}

	f = newFake(t)
	if code := run(f.e, []string{"version", "--json"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	var v map[string]string
	if err := json.Unmarshal(f.out.Bytes(), &v); err != nil || v["version"] != "1.2.3-test" {
		t.Fatalf("json %q err %v", f.out.String(), err)
	}
}

func TestHelp(t *testing.T) {
	f := newFake(t)
	if code := run(f.e, []string{"help"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	for _, want := range []string{"Usage:", "offboard", "--profile", "Exit codes"} {
		if !strings.Contains(f.out.String(), want) {
			t.Errorf("help lacks %q:\n%s", want, f.out.String())
		}
	}

	f = newFake(t)
	if code := run(f.e, []string{"help", "offboard"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	if !strings.Contains(f.out.String(), "confirm") || !strings.Contains(f.out.String(), "intune") {
		t.Fatalf("command help:\n%s", f.out.String())
	}

	f = newFake(t)
	if code := run(f.e, []string{"get", "-h"}); code != exitOK || !strings.Contains(f.out.String(), "top") {
		t.Fatalf("get -h: exit %d out %q", code, f.out.String())
	}
}

// Usage mistakes must exit 2 before any tenant connection is attempted.
func TestUsageErrors(t *testing.T) {
	cases := [][]string{
		{},
		{"bogus"},
		{"get"},
		{"get", "/me", "/users"},
		{"get", "/me", "--bogus"},
		{"user"},
		{"signins"},
		{"offboard", "alice@contoso.com"}, // no --confirm
		{"offboard", "alice@contoso.com", "--confirm", "bob@contoso.com", "--block"}, // confirm mismatch
		{"offboard", "alice@contoso.com", "--confirm", "alice@contoso.com"},          // nothing to do
		{"offboard", "alice@contoso.com", "--confirm", "alice@contoso.com", "--intune", "nuke"},
		{"offboard", "alice@contoso.com", "--confirm", "alice@contoso.com", "--backup-chats"},
	}
	for _, args := range cases {
		f := newFake(t, one)
		if code := run(f.e, args); code != exitUsage {
			t.Errorf("%v: exit %d, want 2 (stderr %q)", args, code, f.err.String())
		}
		if len(f.picked) != 0 {
			t.Errorf("%v: connected although the usage was wrong", args)
		}
		if f.err.Len() == 0 {
			t.Errorf("%v: no message on stderr", args)
		}
	}
}

func TestProfileSelection(t *testing.T) {
	// Two profiles, no selector: refuse and list them.
	f := newFake(t, one, two)
	if code := run(f.e, []string{"get", "/me"}); code != exitUsage {
		t.Fatalf("exit %d", code)
	}
	if !strings.Contains(f.err.String(), "Prod") || !strings.Contains(f.err.String(), "Lab") {
		t.Fatalf("stderr should list profiles: %q", f.err.String())
	}

	// By name, case-insensitive, before the command.
	f = newFake(t, one, two)
	if code := run(f.e, []string{"--profile", "lab", "get", "/me"}); code != exitOK || len(f.picked) != 1 || f.picked[0] != "p2" {
		t.Fatalf("exit %d picked %v stderr %q", code, f.picked, f.err.String())
	}

	// By id, after the command.
	f = newFake(t, one, two)
	if code := run(f.e, []string{"get", "/me", "--profile", "p1"}); code != exitOK || len(f.picked) != 1 || f.picked[0] != "p1" {
		t.Fatalf("exit %d picked %v", code, f.picked)
	}

	// Unknown selector.
	f = newFake(t, one, two)
	if code := run(f.e, []string{"--profile", "nope", "get", "/me"}); code != exitUsage {
		t.Fatalf("exit %d", code)
	}

	// Single profile is implicit.
	f = newFake(t, one)
	if code := run(f.e, []string{"get", "/me"}); code != exitOK || len(f.picked) != 1 {
		t.Fatalf("exit %d picked %v", code, f.picked)
	}

	// No profiles at all.
	f = newFake(t)
	if code := run(f.e, []string{"get", "/me"}); code != exitUsage {
		t.Fatalf("exit %d", code)
	}
}

func TestProfilesCommand(t *testing.T) {
	f := newFake(t, one, two)
	if code := run(f.e, []string{"profiles"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	if !strings.Contains(f.out.String(), "p1") || !strings.Contains(f.out.String(), "device_code") {
		t.Fatalf("table:\n%s", f.out.String())
	}

	f = newFake(t, one, two)
	if code := run(f.e, []string{"profiles", "--json"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	var list []secrets.Profile
	if err := json.Unmarshal(f.out.Bytes(), &list); err != nil || len(list) != 2 {
		t.Fatalf("json %q err %v", f.out.String(), err)
	}
}

func TestGet(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/users" {
			t.Errorf("path %q", r.URL.Path)
		}
		q := r.URL.Query()
		if q.Get("$select") != "id" || q.Get("$top") != "5" {
			t.Errorf("query %v", q)
		}
		w.Write([]byte(`{"value":[{"id":"1"}]}`))
	}
	if code := run(f.e, []string{"get", "/users?$select=id", "--top", "5", "--json"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	if strings.Count(strings.TrimSpace(f.out.String()), "\n") != 0 {
		t.Fatalf("--json must be a single line: %q", f.out.String())
	}
	var resp struct{ Value []map[string]string }
	if err := json.Unmarshal(f.out.Bytes(), &resp); err != nil || len(resp.Value) != 1 || resp.Value[0]["id"] != "1" {
		t.Fatalf("output %q err %v", f.out.String(), err)
	}
}

func TestGetAllFollowsNextLink(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/groups":
			w.Write([]byte(`{"value":[{"id":"g1"}],"@odata.nextLink":"` + f.srv.URL + `/page2"}`))
		case "/page2":
			w.Write([]byte(`{"value":[{"id":"g2"}]}`))
		default:
			t.Errorf("unexpected %s", r.URL.Path)
		}
	}
	if code := run(f.e, []string{"get", "/groups", "--all"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	var resp struct{ Value []map[string]string }
	if err := json.Unmarshal(f.out.Bytes(), &resp); err != nil || len(resp.Value) != 2 || resp.Value[1]["id"] != "g2" {
		t.Fatalf("output %q err %v", f.out.String(), err)
	}
	if len(f.calls) != 2 {
		t.Fatalf("calls %v", f.calls)
	}
}

func TestGetGraphErrorExits1(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusForbidden)
		w.Write([]byte(`{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}`))
	}
	if code := run(f.e, []string{"get", "/users"}); code != exitFail {
		t.Fatalf("exit %d", code)
	}
	if !strings.Contains(f.err.String(), "Insufficient privileges") {
		t.Fatalf("stderr %q", f.err.String())
	}
}

func TestUserSnapshot(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users/alice@contoso.com":
			w.Write([]byte(`{"id":"u1","displayName":"Alice Example","userPrincipalName":"alice@contoso.com","accountEnabled":true,"jobTitle":"Engineer"}`))
		case "/users/alice@contoso.com/memberOf":
			w.Write([]byte(`{"value":[{"@odata.type":"#microsoft.graph.group","displayName":"Sales"}]}`))
		case "/users/alice@contoso.com/licenseDetails":
			w.Write([]byte(`{"value":[{"skuPartNumber":"ENTERPRISEPACK"}]}`))
		default:
			t.Errorf("unexpected %s", r.URL.Path)
		}
	}
	if code := run(f.e, []string{"user", "alice@contoso.com"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	for _, want := range []string{"Alice Example", "Engineer", "Groups (1)", "Sales (group)", "ENTERPRISEPACK"} {
		if !strings.Contains(f.out.String(), want) {
			t.Errorf("missing %q in:\n%s", want, f.out.String())
		}
	}

	f2 := newFake(t, one)
	f2.handler = f.handler
	if code := run(f2.e, []string{"user", "alice@contoso.com", "--json"}); code != exitOK {
		t.Fatalf("exit %d", code)
	}
	var snap map[string]json.RawMessage
	if err := json.Unmarshal(f2.out.Bytes(), &snap); err != nil || snap["profile"] == nil || snap["memberOf"] == nil || snap["licenses"] == nil {
		t.Fatalf("json %q err %v", f2.out.String(), err)
	}
}

func TestSignins(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/auditLogs/signIns" {
			t.Errorf("path %q", r.URL.Path)
		}
		filter := r.URL.Query().Get("$filter")
		if !strings.Contains(filter, "userPrincipalName eq 'alice@contoso.com'") || !strings.Contains(filter, "status/errorCode ne 0") || !strings.Contains(filter, "createdDateTime ge") {
			t.Errorf("filter %q", filter)
		}
		if r.URL.Query().Get("$top") != "10" {
			t.Errorf("top %q", r.URL.Query().Get("$top"))
		}
		w.Write([]byte(`{"value":[{"createdDateTime":"2026-09-20T10:00:00Z","appDisplayName":"Outlook","ipAddress":"1.2.3.4","status":{"errorCode":50126,"failureReason":"Invalid username or password"},"location":{"city":"Oslo","countryOrRegion":"NO"}}]}`))
	}
	if code := run(f.e, []string{"signins", "alice@contoso.com", "--failed", "--days", "3", "--top", "10"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	for _, want := range []string{"FAIL 50126", "Outlook", "Oslo, NO", "1 event(s)"} {
		if !strings.Contains(f.out.String(), want) {
			t.Errorf("missing %q in:\n%s", want, f.out.String())
		}
	}
}

func TestOffboardStreamsStepsAndExitCode(t *testing.T) {
	f := newFake(t, one)
	f.handler = func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "PATCH" && r.URL.Path == "/users/alice@contoso.com":
			var body map[string]any
			json.NewDecoder(r.Body).Decode(&body)
			if body["accountEnabled"] != false {
				t.Errorf("block body %v", body)
			}
			w.WriteHeader(http.StatusNoContent)
		case r.Method == "POST" && r.URL.Path == "/users/alice@contoso.com/revokeSignInSessions":
			w.WriteHeader(http.StatusForbidden)
			w.Write([]byte(`{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}`))
		default:
			w.Write([]byte(`{}`))
		}
	}
	args := []string{"offboard", "alice@contoso.com", "--confirm", "alice@contoso.com", "--block", "--revoke"}
	if code := run(f.e, args); code != exitFail {
		t.Fatalf("exit %d (one step failed → 1), stdout %q stderr %q", code, f.out.String(), f.err.String())
	}
	if !strings.Contains(f.err.String(), "[ok  ] Block sign-in") || !strings.Contains(f.err.String(), "[FAIL] Revoke sessions") {
		t.Fatalf("step stream:\n%s", f.err.String())
	}
	if !strings.Contains(f.out.String(), "1 of 2 step(s) failed") {
		t.Fatalf("summary %q", f.out.String())
	}
	var sawPatch bool
	for _, c := range f.calls {
		if c == "PATCH /users/alice@contoso.com" {
			sawPatch = true
		}
	}
	if !sawPatch {
		t.Fatalf("no block PATCH in %v", f.calls)
	}

	// All steps ok → exit 0; --json prints the PlaybookResult.
	f = newFake(t, one)
	if code := run(f.e, append(args[:len(args)-1], "--json")); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	var res struct {
		OK    bool
		Steps []struct{ Name string }
	}
	if err := json.Unmarshal(f.out.Bytes(), &res); err != nil || !res.OK || len(res.Steps) != 1 || res.Steps[0].Name != "Block sign-in" {
		t.Fatalf("json %q err %v", f.out.String(), err)
	}
}

func TestOffboardReadOnlyIsBlocked(t *testing.T) {
	f := newFake(t, one)
	code := run(f.e, []string{"offboard", "alice@contoso.com", "--confirm", "alice@contoso.com", "--block", "--read-only"})
	if code != exitFail || !f.readOnly {
		t.Fatalf("exit %d readOnly %v", code, f.readOnly)
	}
	if !strings.Contains(f.err.String(), "read-only") {
		t.Fatalf("stderr %q", f.err.String())
	}
	for _, c := range f.calls {
		if strings.HasPrefix(c, "PATCH") || strings.HasPrefix(c, "POST") {
			t.Fatalf("write sent in read-only mode: %v", f.calls)
		}
	}
}

func TestConnectFailureExits1(t *testing.T) {
	f := newFake(t, one)
	f.e.connect = func(context.Context, secrets.Profile, bool, io.Writer) (*session.Session, error) {
		return nil, errors.New("secret not found in keychain")
	}
	if code := run(f.e, []string{"get", "/me"}); code != exitFail {
		t.Fatalf("exit %d", code)
	}
	if !strings.Contains(f.err.String(), "keychain") {
		t.Fatalf("stderr %q", f.err.String())
	}
}
