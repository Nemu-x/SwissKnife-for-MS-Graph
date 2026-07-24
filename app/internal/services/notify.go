package services

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"swissknife-app/internal/session"
)

// NotifyService posts playbook summaries to a Microsoft Teams channel via an
// incoming webhook (a Power Automate "When a Teams webhook request is
// received" flow URL). Webhook posting needs no Graph permission and no
// protected-API approval — it is a plain HTTPS POST with an Adaptive Card.
type NotifyService struct {
	s *session.Session
}

func NewNotifyService(s *session.Session) *NotifyService { return &NotifyService{s: s} }

// NotifyConfig is persisted as notify.json in the app config directory. The
// webhook URL lets anyone post to the channel, so the file is user-scoped
// (0600) and the URL is never written to logs.
type NotifyConfig struct {
	WebhookURL      string `json:"webhookUrl"`
	NotifyPlaybooks bool   `json:"notifyPlaybooks"`
}

func notifyConfigPath(dir string) string { return filepath.Join(dir, "notify.json") }

func loadNotifyConfig(dir string) *NotifyConfig {
	cfg := &NotifyConfig{}
	if dir == "" {
		return cfg
	}
	b, err := os.ReadFile(notifyConfigPath(dir))
	if err != nil {
		return cfg
	}
	_ = json.Unmarshal(b, cfg)
	return cfg
}

// Get returns the saved notification settings.
func (n *NotifyService) Get() (*NotifyConfig, error) {
	return loadNotifyConfig(n.s.ConfigDir()), nil
}

// Set validates and persists the notification settings.
func (n *NotifyService) Set(cfg NotifyConfig) error {
	cfg.WebhookURL = strings.TrimSpace(cfg.WebhookURL)
	if cfg.WebhookURL != "" {
		u, err := url.Parse(cfg.WebhookURL)
		if err != nil || u.Scheme != "https" || u.Host == "" {
			return errors.New("the webhook URL must be a valid https:// address")
		}
	}
	dir := n.s.ConfigDir()
	if dir == "" {
		return errors.New("config directory unavailable")
	}
	b, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(notifyConfigPath(dir), b, 0o600)
}

// Test sends a test card to the configured webhook.
func (n *NotifyService) Test() error {
	cfg := loadNotifyConfig(n.s.ConfigDir())
	if cfg.WebhookURL == "" {
		return errors.New("no webhook URL configured")
	}
	return postAdaptiveCard(context.Background(), cfg.WebhookURL, "SwissKnife test notification",
		[][2]string{{"Status", "Webhook is working"}}, nil)
}

// notifyPlaybookSummary posts a run summary card, best-effort: failures are
// audited, never surfaced to the operator's run result.
func notifyPlaybookSummary(s *session.Session, kind, upn string, steps []Step, canceled bool) {
	cfg := loadNotifyConfig(s.ConfigDir())
	if !cfg.NotifyPlaybooks || cfg.WebhookURL == "" {
		return
	}
	failed := 0
	var failedLines []string
	for _, st := range steps {
		if !st.OK {
			failed++
			line := st.Name
			if st.Error != "" {
				line += ": " + st.Error
			}
			failedLines = append(failedLines, "✗ "+line)
		}
	}
	label := kind
	if label != "" {
		label = strings.ToUpper(label[:1]) + label[1:]
	}
	title := fmt.Sprintf("%s completed — %s", label, upn)
	status := "OK"
	if canceled {
		status = "Canceled"
	} else if failed > 0 {
		status = fmt.Sprintf("%d step(s) failed", failed)
	}
	facts := [][2]string{
		{"User", upn},
		{"Steps", itoa(len(steps))},
		{"Result", status},
	}
	err := postAdaptiveCard(context.Background(), cfg.WebhookURL, title, facts, failedLines)
	s.Record("notify.teams", upn, "kind="+kind, err)
}

// postAdaptiveCard sends a Teams message payload with one Adaptive Card.
func postAdaptiveCard(ctx context.Context, webhookURL, title string, facts [][2]string, extraLines []string) error {
	body := []map[string]any{
		{"type": "TextBlock", "size": "Medium", "weight": "Bolder", "text": title, "wrap": true},
	}
	if len(facts) > 0 {
		fs := make([]map[string]string, 0, len(facts))
		for _, f := range facts {
			fs = append(fs, map[string]string{"title": f[0], "value": f[1]})
		}
		body = append(body, map[string]any{"type": "FactSet", "facts": fs})
	}
	for _, line := range extraLines {
		body = append(body, map[string]any{"type": "TextBlock", "text": line, "wrap": true, "spacing": "None"})
	}
	payload := map[string]any{
		"type": "message",
		"attachments": []any{map[string]any{
			"contentType": "application/vnd.microsoft.card.adaptive",
			"content": map[string]any{
				"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
				"type":    "AdaptiveCard",
				"version": "1.4",
				"body":    body,
			},
		}},
	}
	raw, err := json.Marshal(payload)
	if err != nil {
		return err
	}
	ctx, cancel := context.WithTimeout(ctx, 10*time.Second)
	defer cancel()
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, webhookURL, bytes.NewReader(raw))
	if err != nil {
		return err
	}
	req.Header.Set("Content-Type", "application/json")
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return fmt.Errorf("webhook answered HTTP %d", resp.StatusCode)
	}
	return nil
}
