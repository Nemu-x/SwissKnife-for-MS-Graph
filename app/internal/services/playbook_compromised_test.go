package services

import (
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"unicode"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/pwsh"
)

func TestCompromisedLocksOutThenCleansUpAndKeepsThePasswordOutOfLogs(t *testing.T) {
	var patches []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		p := r.URL.Path
		b, _ := io.ReadAll(r.Body)
		switch {
		case r.Method == "PATCH":
			patches = append(patches, p+" "+string(b))
			w.WriteHeader(http.StatusNoContent)
		case p == "/domains":
			w.Write([]byte(`{"value":[{"id":"contoso.com"}]}`))
		case strings.HasSuffix(p, "/mailFolders/inbox/messageRules"):
			w.Write([]byte(`{"value":[
				{"id":"r1","displayName":".","isEnabled":true,"actions":{"forwardTo":[{"emailAddress":{"address":"x@evil.example"}}]}},
				{"id":"r2","displayName":"team","isEnabled":true,"actions":{"forwardTo":[{"emailAddress":{"address":"boss@contoso.com"}}]}}]}`))
		case strings.Contains(p, "/mailFolders/"):
			w.Write([]byte(`{"id":"F"}`))
		case strings.HasSuffix(p, "/authentication/methods"):
			w.Write([]byte(`{"value":[]}`))
		case p == "/auditLogs/signIns":
			w.Write([]byte(`{"value":[{"createdDateTime":"2026-10-07T03:00:00Z","appDisplayName":"Office","ipAddress":"203.0.113.9","location":{"city":"Lagos","countryOrRegion":"NG"},"status":{"errorCode":0}}]}`))
		case strings.HasPrefix(p, "/users/"):
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","accountEnabled":true}`))
		default:
			w.Write([]byte(`{}`))
		}
	})
	dir := t.TempDir()
	sess.SetJournal(journal.New(filepath.Join(dir, "runs")))
	ps := &recordingPS{}
	e := engine.New(sess, engine.GraphProvider{}, readyProvider{pwsh.BackendExchangePS})
	e.PS = ps
	e.Register(actions.Builtin()...)
	engines.Store(sess, e)
	t.Cleanup(func() { engines.Delete(sess) })

	res, err := NewPlaybookService(sess).Compromised(CompromisedRequest{
		Upn: "ann@contoso.com", Confirm: "ann@contoso.com",
		ResetPassword: true, ResetMfa: true, ClearForwarding: true, DisableRules: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	var names []string
	for _, s := range res.Steps {
		names = append(names, s.Name)
		if !s.OK {
			t.Errorf("step %s failed: %s", s.Name, s.Error)
		}
	}
	want := "Block sign-in,Revoke sessions,Reset password,Reset MFA,Clear mailbox forwarding,Disable suspicious inbox rules"
	if strings.Join(names, ",") != want {
		t.Fatalf("steps %v", names)
	}
	if len(res.TempPassword) != 20 || !strings.ContainsFunc(res.TempPassword, unicode.IsDigit) {
		t.Fatalf("temp password %q", res.TempPassword)
	}
	if len(res.SignIns) != 1 || res.SignIns[0].Location != "Lagos, NG" {
		t.Fatalf("sign-ins %+v", res.SignIns)
	}
	disabled := false
	for _, p := range patches {
		if strings.Contains(p, "/messageRules/r2") {
			t.Errorf("an internal forward must stay: %s", p)
		}
		if strings.Contains(p, "/messageRules/r1") && strings.Contains(p, `"isEnabled":false`) {
			disabled = true
		}
	}
	if !disabled {
		t.Fatalf("the external-forward rule must be disabled: %v", patches)
	}
	if !strings.Contains(strings.Join(ps.calls, "\n"), "Set-Mailbox") {
		t.Fatalf("forwarding must be cleared through Exchange: %v", ps.calls)
	}

	// The temporary password is shown once — never written to disk.
	_ = filepath.WalkDir(dir, func(path string, d os.DirEntry, err error) error {
		if err == nil && !d.IsDir() {
			if b, _ := os.ReadFile(path); strings.Contains(string(b), res.TempPassword) {
				t.Errorf("the temporary password leaked into %s", path)
			}
		}
		return nil
	})
}

func TestCompromisedNeedsTheTypedConfirmation(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) })
	if _, err := NewPlaybookService(sess).Compromised(CompromisedRequest{Upn: "ann@contoso.com", Confirm: "bob@contoso.com"}); err == nil {
		t.Fatal("must refuse without the typed confirmation")
	}
}
