package cli

import (
	"net/http"
	"strings"
	"testing"
)

func signInHandler(writes *[]string) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		if r.Method != "GET" {
			*writes = append(*writes, r.Method+" "+r.URL.Path)
			w.WriteHeader(http.StatusNoContent)
			return
		}
		w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","accountEnabled":true}`))
	}
}

func TestActionPreviewDoesNotWrite(t *testing.T) {
	f := newFake(t, one)
	var writes []string
	f.handler = signInHandler(&writes)
	if code := run(f.e, []string{"action", "user.signIn", "user=ann@contoso.com", "state=blocked"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	out := f.out.String()
	if !strings.Contains(out, "signIn: allowed → blocked") || !strings.Contains(out, "--apply") {
		t.Fatalf("output %q", out)
	}
	if len(writes) != 0 {
		t.Fatalf("preview wrote: %v", writes)
	}
}

func TestActionApply(t *testing.T) {
	f := newFake(t, one)
	var writes []string
	f.handler = signInHandler(&writes)
	if code := run(f.e, []string{"action", "user.signIn", "user=ann@contoso.com", "--apply"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	if len(writes) != 1 || writes[0] != "PATCH /users/u1" {
		t.Fatalf("writes %v", writes)
	}
	if !strings.Contains(f.out.String(), "Applied 1, unchanged 0, failed 0.") {
		t.Fatalf("output %q", f.out.String())
	}
}

func TestActionDestructiveNeedsConfirm(t *testing.T) {
	f := newFake(t, one)
	var writes []string
	f.handler = signInHandler(&writes)
	if code := run(f.e, []string{"action", "user.revokeSessions", "user=ann@contoso.com", "--apply"}); code != exitFail {
		t.Fatalf("exit %d, want %d", code, exitFail)
	}
	if len(writes) != 0 {
		t.Fatalf("writes %v", writes)
	}
}

func TestActionUsageErrors(t *testing.T) {
	f := newFake(t, one)
	if code := run(f.e, []string{"action"}); code != exitUsage {
		t.Fatalf("no args: exit %d", code)
	}
	if code := run(f.e, []string{"action", "user.signIn", "oops"}); code != exitUsage {
		t.Fatalf("bad pair: exit %d", code)
	}
}

func TestActionList(t *testing.T) {
	f := newFake(t, one)
	if code := run(f.e, []string{"action", "list"}); code != exitOK {
		t.Fatalf("exit %d stderr %q", code, f.err.String())
	}
	if !strings.Contains(f.out.String(), "user.signIn") || !strings.Contains(f.out.String(), "state=blocked|allowed") {
		t.Fatalf("output %q", f.out.String())
	}
}
