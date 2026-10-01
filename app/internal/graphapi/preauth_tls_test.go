package graphapi

import (
	"context"
	"strings"
	"testing"
)

// The pre-authorized URL carries its token in the query string, so anything
// but TLS (or loopback, for tests) must be refused before a byte is sent.
func TestPostPreauthorizedRefusesPlainHTTP(t *testing.T) {
	c := New(StaticToken("t"))
	for _, u := range []string{"http://outlook.office.com/import?token=x", "http://10.0.0.5/x", "ftp://x/y", "/relative"} {
		err := c.PostPreauthorized(context.Background(), u, nil, nil)
		if err == nil {
			t.Errorf("%s: expected a refusal", u)
			continue
		}
		if strings.HasPrefix(u, "http://") && !strings.Contains(err.Error(), "https") {
			t.Errorf("%s: error should say https is required, got %v", u, err)
		}
	}
}

func TestIsLoopback(t *testing.T) {
	for host, want := range map[string]bool{"localhost": true, "127.0.0.1": true, "::1": true, "10.0.0.5": false, "example.com": false, "": false} {
		if got := isLoopback(host); got != want {
			t.Errorf("isLoopback(%q) = %v, want %v", host, got, want)
		}
	}
}
