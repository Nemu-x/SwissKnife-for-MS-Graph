package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
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

// purgeFake plays the eDiscovery API. estimates answers successive estimate
// operations; status overrides the operation status; the search query can be
// tampered with to simulate an edit in Purview.
type purgeFake struct {
	srv         *httptest.Server
	estimates   []int64
	status      string
	query       string
	tamper      bool
	searchFails bool
	purges      []string
	preferHdr   string
	caseDeleted bool
	estimated   int
}

func newPurgeFake(t *testing.T, estimates ...int64) (*purgeFake, *engine.Engine, *session.Session) {
	t.Helper()
	old, oldSettle := pollEvery, settle
	pollEvery, settle = time.Millisecond, 0
	t.Cleanup(func() { pollEvery, settle = old, oldSettle })

	f := &purgeFake{estimates: estimates, status: "succeeded"}
	f.srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		switch {
		case r.URL.Path == "/domains":
			w.Write([]byte(`{"value":[{"id":"contoso.com"},{"id":"contoso.onmicrosoft.com"}]}`))
		case r.Method == "POST" && r.URL.Path == "/security/cases/ediscoveryCases":
			w.Write([]byte(`{"id":"c1"}`))
		case r.Method == "DELETE" && r.URL.Path == "/security/cases/ediscoveryCases/c1":
			f.caseDeleted = true
			w.WriteHeader(http.StatusNoContent)
		case r.Method == "POST" && r.URL.Path == "/security/cases/ediscoveryCases/c1/searches":
			if f.searchFails {
				w.WriteHeader(http.StatusForbidden)
				w.Write([]byte(`{"error":{"code":"Forbidden","message":"no role"}}`))
				return
			}
			var body struct{ ContentQuery string }
			_ = json.Unmarshal(b, &body)
			f.query = body.ContentQuery
			w.Write([]byte(`{"id":"s1"}`))
		case r.Method == "GET" && r.URL.Path == "/security/cases/ediscoveryCases/c1/searches/s1":
			q := f.query
			if f.tamper {
				q = `from:"*"`
			}
			b, _ := json.Marshal(map[string]string{"contentQuery": q, "dataSourceScopes": "allTenantMailboxes"})
			w.Write(b)
		case r.Method == "POST" && strings.HasSuffix(r.URL.Path, "/estimateStatistics"):
			w.Header().Set("Location", f.srv.URL+"/ops/estimate")
			w.WriteHeader(http.StatusAccepted)
		case r.URL.Path == "/ops/estimate":
			n := f.estimates[min(f.estimated, len(f.estimates)-1)]
			f.estimated++
			fmt.Fprintf(w, `{"status":%q,"mailboxCount":3,"indexedItemCount":%d}`, f.status, n)
		case r.Method == "POST" && strings.HasSuffix(r.URL.Path, "/purgeData"):
			var body struct{ PurgeType string }
			_ = json.Unmarshal(b, &body)
			f.purges = append(f.purges, body.PurgeType)
			f.preferHdr = r.Header.Get("Prefer")
			w.Header().Set("Location", f.srv.URL+"/ops/purge")
			w.WriteHeader(http.StatusAccepted)
		case r.URL.Path == "/ops/purge":
			fmt.Fprintf(w, `{"status":%q}`, f.status)
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL.Path)
		}
	}))
	t.Cleanup(f.srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(f.srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{})
	e.Register(Builtin()...)
	return f, e, s
}

func TestPurgeConfirmationNamesSenderAndCount(t *testing.T) {
	f, e, _ := newPurgeFake(t, 12)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": " Evil@Phish.example ", "subject": "Invoice  urgent", "since": "2026-10-01"})
	if err != nil {
		t.Fatal(err)
	}
	if f.query != `from:"evil@phish.example" AND subject:"Invoice urgent" AND received>=2026-10-01` {
		t.Fatalf("query %q", f.query)
	}
	c := p.Changes[0]
	if c.Op != "remove" || c.Before != "12 / 3" || c.Target != f.query || c.Note != "purge.recoverable" {
		t.Fatalf("change %+v", c)
	}
	if p.ConfirmTarget != "evil@phish.example 12" {
		t.Fatalf("confirm %q", p.ConfirmTarget)
	}
	if _, err := e.Apply(p.ID, "evil@phish.example"); err == nil {
		t.Fatal("the sender alone must not confirm a purge")
	}
	if _, err := e.Apply(p.ID, "evil@phish.example 12"); err != nil {
		t.Fatal(err)
	}
	if len(f.purges) != 1 || f.purges[0] != "recoverable" {
		t.Fatalf("a recoverable purge runs once: %v", f.purges)
	}
}

func TestPurgeRefusesBroadOrUnsafeQueries(t *testing.T) {
	internal := map[string]bool{"contoso.com": true}
	for _, in := range []engine.Inputs{
		{"sender": "contoso.com", "subject": "x"},   // the tenant's own domain
		{"sender": "gmail.com", "subject": "x"},     // a free-mail provider
		{"sender": "phish.example"},                 // a bare domain without narrowing
		{"sender": `a@b.com" OR from:"*`},           // quote injection
		{"sender": "*@phish.example"},               // wildcard
		{"sender": "a“b@phish.example"},             // unicode quote
		{"sender": "a@b.com", "subject": `x" OR "`}, // quote in the subject
		{"sender": "a@b.com", "subject": "pay (now)"},
		{"sender": "a@b.com", "subject": "x:*"},
		{"sender": "a@b.com", "since": "2026-13-45"},
		{"sender": "a@b.com", "since": "2999-01-01"},
		{"sender": "ceo@contoso.com"}, // a colleague's whole mailbox
		{"sender": "onmicrosoft.com", "subject": "x"},
	} {
		if q, err := purgeQuery(in, internal); err == nil {
			t.Errorf("purgeQuery(%v) = %q, want a refusal", in, q)
		}
	}
	if _, err := purgeQuery(engine.Inputs{"sender": "phish.example", "since": "2026-10-01"}, internal); err != nil {
		t.Errorf("a narrowed external domain is fine: %v", err)
	}
}

func TestPurgePreviewRespectsReadOnlyAndCleansUpOnFailure(t *testing.T) {
	f, e, s := newPurgeFake(t, 1)
	s.SetReadOnly(true)
	if _, err := e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example"}); !errors.Is(err, session.ErrReadOnly) {
		t.Fatalf("read-only preview: %v", err)
	}
	s.SetReadOnly(false)
	f.searchFails = true
	if _, err := e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example"}); err == nil {
		t.Fatal("the failing search must fail the preview")
	}
	if !f.caseDeleted {
		t.Fatal("a failed preview must delete the case it created")
	}
}

func TestPurgeRefusesASearchEditedAfterPreview(t *testing.T) {
	f, e, _ := newPurgeFake(t, 5)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example"})
	if err != nil {
		t.Fatal(err)
	}
	f.tamper = true
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil {
		t.Fatal(err) // the engine reports failures in the result, not as an error
	}
	if len(f.purges) != 0 {
		t.Fatal("nothing may be purged once the search changed")
	}
}

func TestPartialEstimateIsAnError(t *testing.T) {
	f, e, _ := newPurgeFake(t, 0)
	f.status = "partiallySucceeded"
	if _, err := e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example"}); err == nil {
		t.Fatal("a partial estimate must not read as 'nothing matches'")
	}
}

func TestPermanentPurgeStopsWhenTheCountStopsFalling(t *testing.T) {
	f, e, _ := newPurgeFake(t, 250, 150, 0)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example", "purgeType": "permanent"})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil {
		t.Fatal(err)
	}
	if len(f.purges) != 2 || f.purges[0] != "permanentlyDelete" || f.preferHdr != "include-unknown-enum-members" {
		t.Fatalf("purges %v prefer %q", f.purges, f.preferHdr)
	}

	// Items on hold: the count never falls; the loop must stop early.
	f, e, _ = newPurgeFake(t, 80, 80, 80)
	p, _ = e.Plan("mail.purge", engine.Inputs{"sender": "a@phish.example", "purgeType": "permanent"})
	r, err := e.Apply(p.ID, p.ConfirmTarget)
	if err != nil {
		t.Fatal(err)
	}
	if r.Failed != 1 || !strings.Contains(r.Outcomes[0].Error, "hold") || len(f.purges) != 2 {
		t.Fatalf("result %+v purges %v", r, f.purges)
	}
}

func TestPurgeWithNoMatchIsNoOp(t *testing.T) {
	f, e, _ := newPurgeFake(t, 0)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "nobody@phish.example"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Op != "none" {
		t.Fatalf("change %+v", p.Changes[0])
	}
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil || len(f.purges) != 0 {
		t.Fatalf("err %v purges %v", err, f.purges)
	}
}
