package graphapi

import (
	"context"
	"fmt"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync/atomic"
	"testing"
)

// A foreign server must never receive a request (and so never the token).
func newForeignServer(t *testing.T) (*httptest.Server, *atomic.Int32) {
	t.Helper()
	var hits atomic.Int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		hits.Add(1)
		fmt.Fprint(w, `{"value":[]}`)
	}))
	t.Cleanup(srv.Close)
	return srv, &hits
}

func TestAbsoluteURLForeignHostRefused(t *testing.T) {
	foreign, hits := newForeignServer(t)
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		fmt.Fprint(w, `{}`)
	}))
	for _, u := range []string{foreign.URL + "/users", strings.ToUpper(foreign.URL[:4]) + foreign.URL[4:] + "/x"} {
		if err := c.Get(context.Background(), u, nil, nil); err == nil || !strings.Contains(err.Error(), "refusing") {
			t.Errorf("Get(%s) err = %v, want refusal", u, err)
		}
		if _, err := c.PostForLocation(context.Background(), u, nil, map[string]any{}); err == nil {
			t.Errorf("PostForLocation(%s) succeeded", u)
		}
	}
	if hits.Load() != 0 {
		t.Fatalf("foreign server got %d requests", hits.Load())
	}
}

func TestAbsoluteURLSchemeDowngradeRefused(t *testing.T) {
	foreign, hits := newForeignServer(t)
	// An https base: a plain-http link to the very same host is refused too.
	c := New(StaticToken("test-token"), WithBaseURL("https://"+strings.TrimPrefix(foreign.URL, "http://")))
	err := c.Get(context.Background(), foreign.URL+"/users", nil, nil)
	if err == nil || !strings.Contains(err.Error(), "refusing") {
		t.Fatalf("err = %v, want refusal", err)
	}
	if hits.Load() != 0 {
		t.Fatalf("http server got %d requests", hits.Load())
	}
}

func TestUserinfoURLRefused(t *testing.T) {
	c, srv := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		t.Error("request must not be sent")
	}))
	u := strings.Replace(srv.URL, "://", "://user:pw@", 1) + "/users"
	if err := c.Get(context.Background(), u, nil, nil); err == nil {
		t.Fatal("userinfo URL accepted")
	}
}

func TestListAllForeignNextLinkStops(t *testing.T) {
	foreign, hits := newForeignServer(t)
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		fmt.Fprintf(w, `{"value":[{"n":1}],"@odata.nextLink":"%s/items?page=2"}`, foreign.URL)
	}))
	if _, err := c.ListAll(context.Background(), "/items", nil, 0); err == nil {
		t.Fatal("ListAll followed a foreign nextLink")
	}
	if hits.Load() != 0 {
		t.Fatalf("foreign server got %d requests", hits.Load())
	}
}

func TestBetaSameHostAllowed(t *testing.T) {
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/beta/things" {
			t.Errorf("path = %q", r.URL.Path)
		}
		fmt.Fprint(w, `{}`)
	}))
	if err := c.Get(context.Background(), c.Beta("/things"), nil, nil); err != nil {
		t.Fatal(err)
	}
}
