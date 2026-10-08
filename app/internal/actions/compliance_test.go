package actions

import (
	"encoding/json"
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

// purgeFake plays the eDiscovery API: estimates answer with the counts in
// remaining (one per estimate call), purges record their type.
type purgeFake struct {
	srv       *httptest.Server
	remaining []int64
	purges    []string
	query     string
}

func newPurgeFake(t *testing.T, remaining ...int64) (*purgeFake, *engine.Engine) {
	t.Helper()
	old := pollEvery
	pollEvery = time.Millisecond
	t.Cleanup(func() { pollEvery = old })

	f := &purgeFake{remaining: remaining}
	estimates := 0
	f.srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		switch {
		case r.Method == "POST" && r.URL.Path == "/security/cases/ediscoveryCases":
			w.Write([]byte(`{"id":"c1"}`))
		case r.Method == "POST" && r.URL.Path == "/security/cases/ediscoveryCases/c1/searches":
			var body struct{ ContentQuery string }
			_ = json.Unmarshal(b, &body)
			f.query = body.ContentQuery
			w.Write([]byte(`{"id":"s1"}`))
		case r.Method == "POST" && strings.HasSuffix(r.URL.Path, "/estimateStatistics"):
			w.Header().Set("Location", f.srv.URL+"/ops/estimate")
			w.WriteHeader(http.StatusAccepted)
		case r.URL.Path == "/ops/estimate":
			n := f.remaining[min(estimates/2, len(f.remaining)-1)] // two GETs per finished poll
			estimates++
			w.Write([]byte(`{"status":"succeeded","mailboxCount":3,"indexedItemCount":` + itoa(n) + `}`))
		case r.Method == "POST" && strings.HasSuffix(r.URL.Path, "/purgeData"):
			var body struct{ PurgeType string }
			_ = json.Unmarshal(b, &body)
			f.purges = append(f.purges, body.PurgeType)
			w.Header().Set("Location", f.srv.URL+"/ops/purge")
			w.WriteHeader(http.StatusAccepted)
		case r.URL.Path == "/ops/purge":
			w.Write([]byte(`{"status":"succeeded"}`))
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL.Path)
		}
	}))
	t.Cleanup(f.srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(f.srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{})
	e.Register(Builtin()...)
	return f, e
}

func itoa(n int64) string { b, _ := json.Marshal(n); return string(b) }

func TestPurgePreviewsCountsAndNeedsTheSender(t *testing.T) {
	f, e := newPurgeFake(t, 12)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "evil@phish.example", "subject": `Invoice "urgent" (pay)`, "since": "2026-10-01"})
	if err != nil {
		t.Fatal(err)
	}
	if f.query != `from:"evil@phish.example" AND subject:"Invoice urgent pay" AND received>=2026-10-01` {
		t.Fatalf("query %q", f.query)
	}
	c := p.Changes[0]
	if c.Op != "remove" || c.Before != "12 / 3" || c.Target != f.query || c.Note != "purge.recoverable" {
		t.Fatalf("change %+v", c)
	}
	if _, err := e.Apply(p.ID, "someone@else.example"); err == nil {
		t.Fatal("a purge must be confirmed by retyping the sender")
	}
	if len(f.purges) != 0 {
		t.Fatal("nothing may be purged before the confirmation")
	}
	if _, err := e.Apply(p.ID, "evil@phish.example"); err != nil {
		t.Fatal(err)
	}
	if len(f.purges) != 1 || f.purges[0] != "recoverable" {
		t.Fatalf("a recoverable purge runs once: %v", f.purges)
	}
}

func TestPermanentPurgeRepeatsUntilNothingMatches(t *testing.T) {
	f, e := newPurgeFake(t, 250, 150, 0)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "phish.example", "purgeType": "permanent"})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := e.Apply(p.ID, "phish.example"); err != nil {
		t.Fatal(err)
	}
	if len(f.purges) != 2 || f.purges[0] != "permanentlyDelete" {
		t.Fatalf("purges %v, want two permanent runs", f.purges)
	}
}

func TestPurgeWithNoMatchIsNoOp(t *testing.T) {
	f, e := newPurgeFake(t, 0)
	p, err := e.Plan("mail.purge", engine.Inputs{"sender": "nobody@phish.example"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Op != "none" {
		t.Fatalf("change %+v", p.Changes[0])
	}
	if _, err := e.Apply(p.ID, "nobody@phish.example"); err != nil || len(f.purges) != 0 {
		t.Fatalf("err %v purges %v", err, f.purges)
	}
}

func TestPurgeQueryRefusesWideningInput(t *testing.T) {
	for _, in := range []engine.Inputs{
		{"sender": `a@b.com" OR from:"*`},
		{"sender": "*"},
		{"sender": "a@b.com", "since": "yesterday"},
	} {
		if _, err := purgeQuery(in); err == nil {
			t.Errorf("purgeQuery(%v) must be refused", in)
		}
	}
}
