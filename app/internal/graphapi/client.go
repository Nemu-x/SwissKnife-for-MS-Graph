package graphapi

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"
)

const DefaultBaseURL = "https://graph.microsoft.com/v1.0"

// TokenSource yields a current access token (auto-refresh is the implementation's concern).
// Services never see the token: it lives only inside the client (ADR-002).
type TokenSource interface {
	Token(ctx context.Context) (string, error)
}

// StaticToken is a fixed token (tests).
type StaticToken string

func (s StaticToken) Token(ctx context.Context) (string, error) { return string(s), nil }

type Client struct {
	http    *http.Client
	baseURL string
	tokens  TokenSource

	maxRetries int
	// backoff is a function so tests do not actually sleep
	sleep func(ctx context.Context, d time.Duration) error
}

type Option func(*Client)

func WithBaseURL(u string) Option { return func(c *Client) { c.baseURL = strings.TrimRight(u, "/") } }

func WithHTTPClient(h *http.Client) Option { return func(c *Client) { c.http = h } }

func WithMaxRetries(n int) Option { return func(c *Client) { c.maxRetries = n } }

func New(tokens TokenSource, opts ...Option) *Client {
	// No whole-request timeout: it would kill slow-but-progressing streams
	// (large downloads/uploads). Stalls are caught by ResponseHeaderTimeout
	// here plus the idle watchdog on streaming reads (files.go); operations
	// are cancellable via context.
	transport := http.DefaultTransport.(*http.Transport).Clone()
	transport.ResponseHeaderTimeout = 30 * time.Second
	c := &Client{
		http:       &http.Client{Transport: transport},
		baseURL:    DefaultBaseURL,
		tokens:     tokens,
		maxRetries: 4,
		sleep: func(ctx context.Context, d time.Duration) error {
			t := time.NewTimer(d)
			defer t.Stop()
			select {
			case <-ctx.Done():
				return ctx.Err()
			case <-t.C:
				return nil
			}
		},
	}
	for _, o := range opts {
		o(c)
	}
	return c
}

// Do performs a Graph request. path is relative ("/users") or an absolute URL
// (e.g. @odata.nextLink). The response is decoded into out (if out != nil and a body exists).
func (c *Client) Do(ctx context.Context, method, path string, params url.Values, body any, out any) error {
	raw, _, err := c.doRaw(ctx, method, path, params, body)
	if err != nil {
		return err
	}
	if out != nil && len(raw) > 0 {
		if err := json.Unmarshal(raw, out); err != nil {
			return fmt.Errorf("graph: decode response: %w", err)
		}
	}
	return nil
}

// PostForLocation performs a POST and returns the Location header — Graph
// replies 202 + a monitor URL for long-running operations (driveItem copy).
func (c *Client) PostForLocation(ctx context.Context, path string, params url.Values, body any) (string, error) {
	_, loc, err := c.doRaw(ctx, http.MethodPost, path, params, body)
	return loc, err
}

func (c *Client) doRaw(ctx context.Context, method, path string, params url.Values, body any) ([]byte, string, error) {
	u := path
	if !strings.HasPrefix(path, "http://") && !strings.HasPrefix(path, "https://") {
		u = c.baseURL + "/" + strings.TrimLeft(path, "/")
	}
	if len(params) > 0 {
		sep := "?"
		if strings.Contains(u, "?") {
			sep = "&"
		}
		u += sep + params.Encode()
	}

	var payload []byte
	if body != nil {
		b, err := json.Marshal(body)
		if err != nil {
			return nil, "", fmt.Errorf("graph: encode request: %w", err)
		}
		payload = b
	}

	method = strings.ToUpper(method)

	var raw []byte
	var loc string
	err := c.withRetry(ctx, method, func() (time.Duration, error) {
		var retryAfter time.Duration
		var err error
		raw, loc, retryAfter, err = c.once(ctx, method, u, payload)
		return retryAfter, err
	})
	if err != nil {
		return nil, "", err
	}
	return raw, loc, nil
}

// withRetry runs attempt until it succeeds or retries are exhausted, honoring
// Retry-After (returned by attempt) with exponential backoff otherwise. It is
// shared by the JSON paths (doRaw) and the streaming transfer paths (files.go)
// so throttling is handled uniformly.
func (c *Client) withRetry(ctx context.Context, method string, attempt func() (time.Duration, error)) error {
	for try := 0; ; try++ {
		retryAfter, err := attempt()
		if err == nil {
			return nil
		}
		if try >= c.maxRetries || !retryable(method, err) {
			return err
		}
		delay := retryAfter
		if delay <= 0 {
			delay = time.Second << try // 1s, 2s, 4s, 8s
		}
		// Graph occasionally answers with multi-minute Retry-After values;
		// honor the signal but keep a single operation from stalling forever.
		const maxRetryAfter = 60 * time.Second
		if delay > maxRetryAfter {
			delay = maxRetryAfter
		}
		if serr := c.sleep(ctx, delay); serr != nil {
			return serr
		}
	}
}

// retryable: 429 (throttling) is retried for any method — Graph did not process
// the request; 503/504 and transport-level failures (connection reset, EOF) —
// idempotent methods only. Context cancellation is never retried.
func retryable(method string, err error) bool {
	if errors.Is(err, context.Canceled) || errors.Is(err, context.DeadlineExceeded) {
		return false
	}
	idempotent := method == http.MethodGet || method == http.MethodPut || method == http.MethodDelete
	var ge *GraphError
	if !errors.As(err, &ge) {
		return idempotent // transport error: the request may or may not have been processed
	}
	switch ge.StatusCode {
	case 429:
		return true
	case 503, 504:
		return idempotent
	}
	return false
}

func (c *Client) once(ctx context.Context, method, u string, payload []byte) ([]byte, string, time.Duration, error) {
	var bodyReader io.Reader
	if payload != nil {
		bodyReader = bytes.NewReader(payload)
	}
	req, err := http.NewRequestWithContext(ctx, method, u, bodyReader)
	if err != nil {
		return nil, "", 0, err
	}

	token, err := c.tokens.Token(ctx)
	if err != nil {
		return nil, "", 0, fmt.Errorf("graph: acquire token: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "application/json")
	if payload != nil {
		req.Header.Set("Content-Type", "application/json")
	}

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", 0, err
	}
	defer resp.Body.Close()

	raw, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, "", 0, err
	}

	if resp.StatusCode >= 200 && resp.StatusCode < 300 {
		loc := resp.Header.Get("Location")
		if resp.StatusCode == http.StatusNoContent {
			return nil, loc, 0, nil
		}
		return raw, loc, 0, nil
	}

	return nil, "", parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, raw)
}

func parseRetryAfter(v string) time.Duration {
	if v == "" {
		return 0
	}
	if secs, err := time.ParseDuration(v + "s"); err == nil {
		return secs
	}
	if t, err := http.ParseTime(v); err == nil {
		if d := time.Until(t); d > 0 {
			return d
		}
	}
	return 0
}

func parseGraphError(resp *http.Response, raw []byte) *GraphError {
	ge := &GraphError{
		StatusCode: resp.StatusCode,
		Code:       http.StatusText(resp.StatusCode),
		RequestID:  resp.Header.Get("request-id"),
	}
	if resp.Request != nil && resp.Request.URL != nil {
		ge.Path = resp.Request.URL.Path
	}
	var envelope struct {
		Error struct {
			Code       string `json:"code"`
			Message    string `json:"message"`
			InnerError struct {
				RequestID string `json:"request-id"`
			} `json:"innerError"`
		} `json:"error"`
	}
	if err := json.Unmarshal(raw, &envelope); err == nil && envelope.Error.Code != "" {
		ge.Code = envelope.Error.Code
		ge.Message = envelope.Error.Message
		if ge.RequestID == "" {
			ge.RequestID = envelope.Error.InnerError.RequestID
		}
	} else {
		// non-JSON body: do not carry a potentially sensitive dump in full
		const maxSnippet = 200
		s := string(raw)
		if len(s) > maxSnippet {
			s = s[:maxSnippet] + "..."
		}
		ge.Message = s
	}
	return ge
}

// Convenience shortcuts.

func (c *Client) Get(ctx context.Context, path string, params url.Values, out any) error {
	return c.Do(ctx, http.MethodGet, path, params, nil, out)
}

func (c *Client) Post(ctx context.Context, path string, body, out any) error {
	return c.Do(ctx, http.MethodPost, path, nil, body, out)
}

func (c *Client) Patch(ctx context.Context, path string, body, out any) error {
	return c.Do(ctx, http.MethodPatch, path, nil, body, out)
}

func (c *Client) Put(ctx context.Context, path string, body, out any) error {
	return c.Do(ctx, http.MethodPut, path, nil, body, out)
}

func (c *Client) Delete(ctx context.Context, path string) error {
	return c.Do(ctx, http.MethodDelete, path, nil, nil, nil)
}
