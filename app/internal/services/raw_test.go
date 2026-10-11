package services

import (
	"errors"
	"fmt"
	"net/http"
	"net/http/httptest"
	"sync/atomic"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

func newRawFixture(t *testing.T) (*RawService, *session.Session, *atomic.Int32) {
	t.Helper()
	var writes atomic.Int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			writes.Add(1)
		}
		fmt.Fprint(w, `{"ok":true}`)
	}))
	t.Cleanup(srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "p")
	return NewRawService(s), s, &writes
}

// A profile capped at "write" must not reach action endpoints or $batch
// through the playground: every non-GET needs the destructive ceiling.
func TestRawWritesNeedDestructiveCeiling(t *testing.T) {
	r, s, writes := newRawFixture(t)
	s.SetPolicy("prof", session.Policy{MaxDanger: "write"})
	if _, err := r.Send("GET", "/users", ""); err != nil {
		t.Fatalf("GET must pass: %v", err)
	}
	for _, m := range []string{"POST", "PATCH", "PUT", "DELETE"} {
		_, err := r.Send(m, "/$batch", `{"requests":[]}`)
		if !errors.Is(err, session.ErrPolicyDestructive) {
			t.Errorf("%s err = %v, want the destructive-policy refusal", m, err)
		}
	}
	if writes.Load() != 0 {
		t.Fatalf("%d writes reached the server", writes.Load())
	}

	// Without the cap the same write goes through.
	s.SetPolicy("prof", session.Policy{})
	if _, err := r.Send("POST", "/devices/x/wipe", `{}`); err != nil {
		t.Fatalf("POST without a cap: %v", err)
	}
	if writes.Load() != 1 {
		t.Fatalf("writes = %d, want 1", writes.Load())
	}
}

func TestRawReadOnlyBlocksWrites(t *testing.T) {
	r, s, writes := newRawFixture(t)
	s.SetReadOnly(true)
	if _, err := r.Send("POST", "/users", `{}`); !errors.Is(err, session.ErrReadOnly) {
		t.Fatalf("err = %v, want ErrReadOnly", err)
	}
	if _, err := r.Send("GET", "/users", ""); err != nil {
		t.Fatalf("GET in read-only: %v", err)
	}
	if writes.Load() != 0 {
		t.Fatal("write reached the server in read-only mode")
	}
}
