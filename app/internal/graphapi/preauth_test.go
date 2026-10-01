package graphapi

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"sync/atomic"
	"testing"
)

// A pre-authorized session URL carries its own token: the client must not add
// a Bearer header, must still send JSON, and must decode the reply.
func TestPostPreauthorizedSendsNoBearerAndDecodes(t *testing.T) {
	var gotBody map[string]any
	c, srv := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			t.Errorf("method = %s, want POST", r.Method)
		}
		if got := r.Header.Get("Authorization"); got != "" {
			t.Errorf("Authorization must be absent for a pre-authorized URL, got %q", got)
		}
		if got := r.Header.Get("Content-Type"); got != "application/json" {
			t.Errorf("Content-Type = %q", got)
		}
		if got := r.URL.Query().Get("authtoken"); got != "abc" {
			t.Errorf("session token must travel in the URL untouched, got %q", got)
		}
		b, _ := io.ReadAll(r.Body)
		_ = json.Unmarshal(b, &gotBody)
		fmt.Fprint(w, `{"itemId":"new-1","changeKey":"ck"}`)
	}))

	var out struct {
		ItemID string `json:"itemId"`
	}
	err := c.PostPreauthorized(context.Background(), srv.URL+"/importItem?authtoken=abc",
		map[string]any{"FolderId": "f1", "Mode": "create", "Data": "QUJD"}, &out)
	if err != nil {
		t.Fatal(err)
	}
	if out.ItemID != "new-1" {
		t.Errorf("itemId = %q", out.ItemID)
	}
	if gotBody["FolderId"] != "f1" || gotBody["Mode"] != "create" || gotBody["Data"] != "QUJD" {
		t.Errorf("body = %v", gotBody)
	}
}

// Throttling on the session URL is retried like any Graph call; other errors
// surface as *GraphError so callers can inspect the status.
func TestPostPreauthorizedRetriesThrottlingAndParsesErrors(t *testing.T) {
	var calls atomic.Int32
	c, srv := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch calls.Add(1) {
		case 1:
			w.Header().Set("Retry-After", "1")
			w.WriteHeader(429)
			fmt.Fprint(w, `{"error":{"code":"TooManyRequests","message":"slow down"}}`)
		default:
			w.WriteHeader(401)
			fmt.Fprint(w, `{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}`)
		}
	}))

	err := c.PostPreauthorized(context.Background(), srv.URL+"/importItem?authtoken=x", map[string]any{"a": 1}, nil)
	ge, ok := err.(*GraphError)
	if !ok {
		t.Fatalf("want *GraphError, got %T: %v", err, err)
	}
	if ge.StatusCode != 401 || ge.Code != "InvalidAuthenticationToken" {
		t.Errorf("unexpected error: %+v", ge)
	}
	if calls.Load() != 2 {
		t.Errorf("calls = %d, want 2 (one 429 retry, then the 401 is final)", calls.Load())
	}

	if err := c.PostPreauthorized(context.Background(), "/relative", nil, nil); err == nil {
		t.Error("a relative path must be rejected — the session URL is never built from the Graph base")
	}
}
