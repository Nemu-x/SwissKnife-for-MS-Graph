package services

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"
	"time"
)

func TestNotifyConfigRoundTripAndValidation(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { w.Write([]byte(`{}`)) })
	sess.SetConfigDir(t.TempDir())
	n := NewNotifyService(sess)

	if err := n.Set(NotifyConfig{WebhookURL: "http://insecure.example", NotifyPlaybooks: true}); err == nil {
		t.Fatal("plain http webhook must be rejected")
	}
	if err := n.Set(NotifyConfig{WebhookURL: "https://prod-42.westeurope.logic.azure.com/workflows/x", NotifyPlaybooks: true}); err != nil {
		t.Fatal(err)
	}
	cfg, err := n.Get()
	if err != nil || !cfg.NotifyPlaybooks || !strings.Contains(cfg.WebhookURL, "logic.azure.com") {
		t.Fatalf("round trip failed: %+v %v", cfg, err)
	}
}

// TestOffboardPostsTeamsSummary: a finished playbook must post one Adaptive
// Card with the target UPN to the configured webhook.
func TestOffboardPostsTeamsSummary(t *testing.T) {
	received := make(chan string, 1)
	hook := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		received <- string(b)
		w.WriteHeader(202)
	}))
	t.Cleanup(hook.Close)

	var calls []string
	sess := harness(t, offboardHarness(t, &calls, map[string]string{}))
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	// Config validation requires https; write the test-server URL directly.
	cfgJSON, _ := json.Marshal(NotifyConfig{WebhookURL: hook.URL, NotifyPlaybooks: true})
	if err := os.WriteFile(notifyConfigPath(dir), cfgJSON, 0o600); err != nil {
		t.Fatal(err)
	}

	pb := NewPlaybookService(sess)
	if _, err := pb.Offboard(fullOffboardRequest()); err != nil {
		t.Fatal(err)
	}

	select {
	case body := <-received:
		if !strings.Contains(body, "AdaptiveCard") || !strings.Contains(body, "dep@contoso.com") {
			t.Fatalf("card payload wrong: %s", body)
		}
		if !strings.Contains(body, "Offboard completed") {
			t.Fatalf("title missing: %s", body)
		}
	case <-time.After(3 * time.Second):
		t.Fatal("no webhook POST within 3s")
	}
}

// TestNotifyDisabledPostsNothing: without the toggle no request is sent.
func TestNotifyDisabledPostsNothing(t *testing.T) {
	posted := make(chan struct{}, 1)
	hook := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		select {
		case posted <- struct{}{}:
		default:
		}
	}))
	t.Cleanup(hook.Close)

	var calls []string
	sess := harness(t, offboardHarness(t, &calls, map[string]string{}))
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	cfgJSON, _ := json.Marshal(NotifyConfig{WebhookURL: hook.URL, NotifyPlaybooks: false})
	if err := os.WriteFile(notifyConfigPath(dir), cfgJSON, 0o600); err != nil {
		t.Fatal(err)
	}

	pb := NewPlaybookService(sess)
	if _, err := pb.Offboard(fullOffboardRequest()); err != nil {
		t.Fatal(err)
	}
	select {
	case <-posted:
		t.Fatal("disabled notifications must not post")
	case <-time.After(300 * time.Millisecond): // grace window for a wrong-behaving goroutine
	}
}
