package graphapi

import (
	"context"
	"errors"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

// A pre-authorized upload must never follow a redirect: the token (and on
// 307/308 the whole body) would be re-sent to whatever host the 3xx names.
func TestPostPreauthorizedDoesNotFollowRedirects(t *testing.T) {
	hits := 0
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		hits++
		if r.URL.Path == "/elsewhere" {
			t.Errorf("redirect target was requested")
		}
		http.Redirect(w, r, "/elsewhere", http.StatusTemporaryRedirect)
	}))
	defer srv.Close()
	c := New(StaticToken("t"), WithBaseURL(srv.URL), WithMaxRetries(0))
	err := c.PostPreauthorized(context.Background(), srv.URL+"/import?authtoken=secret", map[string]any{"x": 1}, nil)
	var ge *GraphError
	if !errors.As(err, &ge) || ge.StatusCode != http.StatusTemporaryRedirect {
		t.Fatalf("expected a 307 GraphError, got %v", err)
	}
	if hits != 1 {
		t.Errorf("expected exactly one request, got %d", hits)
	}
}

// Transport errors quote the request URL; the token in its query string must
// not leak into error text that ends up in reports and the UI.
func TestPostPreauthorizedStripsTokenFromTransportErrors(t *testing.T) {
	c := New(StaticToken("t"), WithMaxRetries(0))
	// Loopback is allowed over plain http; port 1 refuses the connection.
	err := c.PostPreauthorized(context.Background(), "http://127.0.0.1:1/import?authtoken=supersecret", nil, nil)
	if err == nil {
		t.Fatal("expected a connection error")
	}
	if strings.Contains(err.Error(), "supersecret") || strings.Contains(err.Error(), "authtoken") {
		t.Errorf("error text leaks the token: %v", err)
	}
	if !strings.Contains(err.Error(), "/import") {
		t.Errorf("error should still name the path: %v", err)
	}
}
