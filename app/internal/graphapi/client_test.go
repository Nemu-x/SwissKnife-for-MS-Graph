package graphapi

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"net/url"
	"sync/atomic"
	"testing"
	"time"
)

func newTestClient(t *testing.T, handler http.Handler) (*Client, *httptest.Server) {
	t.Helper()
	srv := httptest.NewServer(handler)
	t.Cleanup(srv.Close)
	c := New(StaticToken("test-token"), WithBaseURL(srv.URL))
	// мгновенный "сон" в ретраях
	c.sleep = func(ctx context.Context, d time.Duration) error { return nil }
	return c, srv
}

func TestGetDecodesAndSendsAuth(t *testing.T) {
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authorization"); got != "Bearer test-token" {
			t.Errorf("Authorization = %q", got)
		}
		if r.URL.Path != "/users/alice" {
			t.Errorf("path = %q", r.URL.Path)
		}
		fmt.Fprint(w, `{"id":"1","displayName":"Alice"}`)
	}))

	var out struct {
		ID          string `json:"id"`
		DisplayName string `json:"displayName"`
	}
	if err := c.Get(context.Background(), "/users/alice", nil, &out); err != nil {
		t.Fatal(err)
	}
	if out.DisplayName != "Alice" {
		t.Errorf("DisplayName = %q", out.DisplayName)
	}
}

func TestGraphErrorParsing(t *testing.T) {
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("request-id", "req-123")
		w.WriteHeader(403)
		fmt.Fprint(w, `{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}`)
	}))

	err := c.Get(context.Background(), "/users", nil, nil)
	ge, ok := err.(*GraphError)
	if !ok {
		t.Fatalf("want *GraphError, got %T: %v", err, err)
	}
	if ge.Code != "Authorization_RequestDenied" || ge.RequestID != "req-123" || !IsForbidden(err) {
		t.Errorf("unexpected error: %+v", ge)
	}
}

func TestRetryOn429RespectsRetryAfter(t *testing.T) {
	var calls atomic.Int32
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if calls.Add(1) <= 2 {
			w.Header().Set("Retry-After", "1")
			w.WriteHeader(429)
			fmt.Fprint(w, `{"error":{"code":"TooManyRequests","message":"throttled"}}`)
			return
		}
		fmt.Fprint(w, `{"ok":true}`)
	}))

	var slept []time.Duration
	c.sleep = func(ctx context.Context, d time.Duration) error {
		slept = append(slept, d)
		return nil
	}

	// 429 ретраится и для POST (запрос не был обработан)
	if err := c.Post(context.Background(), "/things", map[string]any{"a": 1}, nil); err != nil {
		t.Fatal(err)
	}
	if calls.Load() != 3 {
		t.Errorf("calls = %d, want 3", calls.Load())
	}
	for _, d := range slept {
		if d != time.Second {
			t.Errorf("slept %v, want 1s from Retry-After", d)
		}
	}
}

func TestNoRetryOn503ForPost(t *testing.T) {
	var calls atomic.Int32
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		calls.Add(1)
		w.WriteHeader(503)
		fmt.Fprint(w, `{"error":{"code":"ServiceUnavailable","message":"down"}}`)
	}))

	err := c.Post(context.Background(), "/things", map[string]any{"a": 1}, nil)
	if err == nil {
		t.Fatal("want error")
	}
	if calls.Load() != 1 {
		t.Errorf("calls = %d, want 1 (no retry for POST on 503)", calls.Load())
	}
}

func TestRetryGivesUpAfterMax(t *testing.T) {
	var calls atomic.Int32
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		calls.Add(1)
		w.WriteHeader(429)
		fmt.Fprint(w, `{"error":{"code":"TooManyRequests","message":"throttled"}}`)
	}))
	c.maxRetries = 2

	err := c.Get(context.Background(), "/x", nil, nil)
	if err == nil {
		t.Fatal("want error")
	}
	if calls.Load() != 3 { // 1 + 2 ретрая
		t.Errorf("calls = %d, want 3", calls.Load())
	}
}

func TestListAllFollowsNextLink(t *testing.T) {
	var srvURL string
	mux := http.NewServeMux()
	mux.HandleFunc("/items", func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Query().Get("page") {
		case "":
			fmt.Fprintf(w, `{"value":[{"n":1},{"n":2}],"@odata.nextLink":"%s/items?page=2"}`, srvURL)
		case "2":
			fmt.Fprint(w, `{"value":[{"n":3}]}`)
		}
	})
	c, srv := newTestClient(t, mux)
	srvURL = srv.URL

	type item struct {
		N int `json:"n"`
	}
	items, err := ListAllInto[item](context.Background(), c, "/items", nil, 0)
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 3 || items[2].N != 3 {
		t.Errorf("items = %+v", items)
	}
}

func TestListAllHonorsMaxItems(t *testing.T) {
	var srvURL string
	var pages atomic.Int32
	mux := http.NewServeMux()
	mux.HandleFunc("/items", func(w http.ResponseWriter, r *http.Request) {
		pages.Add(1)
		fmt.Fprintf(w, `{"value":[{"n":1},{"n":2}],"@odata.nextLink":"%s/items?page=next"}`, srvURL)
	})
	c, srv := newTestClient(t, mux)
	srvURL = srv.URL

	items, err := c.ListAll(context.Background(), "/items", nil, 3)
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 3 {
		t.Errorf("len = %d, want 3", len(items))
	}
	if pages.Load() != 2 {
		t.Errorf("pages fetched = %d, want 2", pages.Load())
	}
}

func TestNoContentAndParams(t *testing.T) {
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == http.MethodDelete {
			w.WriteHeader(204)
			return
		}
		if got := r.URL.Query().Get("$top"); got != "5" {
			t.Errorf("$top = %q", got)
		}
		fmt.Fprint(w, `{"value":[]}`)
	}))

	if err := c.Delete(context.Background(), "/items/1"); err != nil {
		t.Fatal(err)
	}
	params := url.Values{"$top": {"5"}}
	var out json.RawMessage
	if err := c.Get(context.Background(), "/items", params, &out); err != nil {
		t.Fatal(err)
	}
}
