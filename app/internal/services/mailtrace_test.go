package services

import (
	"encoding/json"
	"net/http"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/session"
)

func TestMailTraceBuildsFilterAndCapsWindow(t *testing.T) {
	var gotPath, gotFilter, gotTop string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		gotPath = r.URL.Path
		gotFilter = r.URL.Query().Get("$filter")
		gotTop = r.URL.Query().Get("$top")
		w.Write([]byte(`{"value":[{"id":"t1","status":"delivered"}]}`))
	})
	mt := NewMailTraceService(sess)

	rows, err := mt.Trace(TraceQuery{Sender: "o'brien@contoso.com", Recipient: "bob@example.org", Days: 30, Top: 20})
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 1 {
		t.Fatalf("rows: %d", len(rows))
	}
	if gotPath != "/admin/exchange/tracing/messageTraces" {
		t.Errorf("path: %q", gotPath)
	}
	if gotTop != "20" {
		t.Errorf("$top: %q", gotTop)
	}
	for _, want := range []string{
		"senderAddress eq 'o''brien@contoso.com'",
		"recipientAddress eq 'bob@example.org'",
		"receivedDateTime ge ",
		"receivedDateTime le ",
	} {
		if !strings.Contains(gotFilter, want) {
			t.Errorf("filter %q lacks %q", gotFilter, want)
		}
	}
	// The API rejects windows over 10 days; 30 must have been clamped.
	parts := strings.Split(gotFilter, " and ")
	var start, end time.Time
	for _, p := range parts {
		switch {
		case strings.HasPrefix(p, "receivedDateTime ge "):
			start, _ = time.Parse("2006-01-02T15:04:05Z", strings.TrimPrefix(p, "receivedDateTime ge "))
		case strings.HasPrefix(p, "receivedDateTime le "):
			end, _ = time.Parse("2006-01-02T15:04:05Z", strings.TrimPrefix(p, "receivedDateTime le "))
		}
	}
	if start.IsZero() || end.IsZero() {
		t.Fatalf("both bounds must be present: %q", gotFilter)
	}
	if d := end.Sub(start); d > 10*24*time.Hour+time.Minute || d < 10*24*time.Hour-time.Minute {
		t.Errorf("window not clamped to 10 days: %v", d)
	}
}

func TestMailTraceDefaultsAndRequiresAnAddress(t *testing.T) {
	var gotFilter, gotTop string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		gotFilter = r.URL.Query().Get("$filter")
		gotTop = r.URL.Query().Get("$top")
		w.Write([]byte(`{"value":[]}`))
	})
	mt := NewMailTraceService(sess)

	if _, err := mt.Trace(TraceQuery{}); err == nil {
		t.Fatal("empty query must be rejected before hitting Graph")
	}
	if _, err := mt.Trace(TraceQuery{Recipient: "bob@example.org"}); err != nil {
		t.Fatal(err)
	}
	if gotTop != "100" {
		t.Errorf("default $top: %q", gotTop)
	}
	if strings.Contains(gotFilter, "senderAddress") {
		t.Errorf("no sender filter expected: %q", gotFilter)
	}
}

func TestMailTraceDetailsPath(t *testing.T) {
	var gotPath string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		gotPath = r.URL.Path
		w.Write([]byte(`{"value":[{"event":"Receive"},{"event":"Deliver"}]}`))
	})
	mt := NewMailTraceService(sess)

	rows, err := mt.Details("7e3b2b2e-1b5e-4b17-80cc-2af6c1d9a3b1", "robert@contoso.com")
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 2 {
		t.Fatalf("rows: %d", len(rows))
	}
	want := "/admin/exchange/tracing/messageTraces/7e3b2b2e-1b5e-4b17-80cc-2af6c1d9a3b1/getDetailsByRecipient(recipientAddress='robert@contoso.com')"
	if gotPath != want {
		t.Errorf("path:\n got %q\nwant %q", gotPath, want)
	}
	if _, err := mt.Details("", "robert@contoso.com"); err == nil {
		t.Error("missing id must be rejected")
	}
}

func TestMailTracePrerequisiteLooksUpServicePrincipal(t *testing.T) {
	var gotFilter string
	present := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/servicePrincipals" {
			t.Errorf("path: %q", r.URL.Path)
		}
		gotFilter = r.URL.Query().Get("$filter")
		if present {
			w.Write([]byte(`{"value":[{"id":"sp1"}]}`))
		} else {
			w.Write([]byte(`{"value":[]}`))
		}
	})
	mt := NewMailTraceService(sess)

	ok, err := mt.Prerequisite()
	if err != nil {
		t.Fatal(err)
	}
	if ok {
		t.Error("no SP in the tenant → false")
	}
	if gotFilter != "appId eq '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373'" {
		t.Errorf("filter: %q", gotFilter)
	}
	present = true
	if ok, err = mt.Prerequisite(); err != nil || !ok {
		t.Errorf("SP present → true, got %v %v", ok, err)
	}
}

func TestMailTraceProvisionPostsServicePrincipal(t *testing.T) {
	var gotMethod, gotPath string
	var body map[string]string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		gotMethod, gotPath = r.Method, r.URL.Path
		json.NewDecoder(r.Body).Decode(&body)
		w.WriteHeader(http.StatusCreated)
		w.Write([]byte(`{"id":"sp1","appId":"8bd644d1-64a1-4d4b-ae52-2e0cbf64e373"}`))
	})
	mt := NewMailTraceService(sess)

	if err := mt.Provision(); err != nil {
		t.Fatal(err)
	}
	if gotMethod != http.MethodPost || gotPath != "/servicePrincipals" {
		t.Errorf("got %s %s", gotMethod, gotPath)
	}
	if body["appId"] != "8bd644d1-64a1-4d4b-ae52-2e0cbf64e373" {
		t.Errorf("body: %v", body)
	}
}

func TestMailTraceProvisionBlockedInReadOnly(t *testing.T) {
	called := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { called = true })
	sess.SetReadOnly(true)
	mt := NewMailTraceService(sess)

	if err := mt.Provision(); err != session.ErrReadOnly {
		t.Fatalf("want ErrReadOnly, got %v", err)
	}
	if called {
		t.Error("write reached the server despite read-only mode")
	}
}
