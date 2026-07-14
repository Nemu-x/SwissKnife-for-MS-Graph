package graphapi

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"
)

const DefaultBaseURL = "https://graph.microsoft.com/v1.0"

// TokenSource выдаёт актуальный access token (авто-refresh — забота реализации).
// Сервисы токен не видят: он живёт только внутри клиента (ADR-002).
type TokenSource interface {
	Token(ctx context.Context) (string, error)
}

// StaticToken — фиксированный токен (тесты).
type StaticToken string

func (s StaticToken) Token(ctx context.Context) (string, error) { return string(s), nil }

type Client struct {
	http    *http.Client
	baseURL string
	tokens  TokenSource

	maxRetries int
	// backoff задаётся функцией, чтобы тесты не спали по-настоящему
	sleep func(ctx context.Context, d time.Duration) error
}

type Option func(*Client)

func WithBaseURL(u string) Option { return func(c *Client) { c.baseURL = strings.TrimRight(u, "/") } }

func WithHTTPClient(h *http.Client) Option { return func(c *Client) { c.http = h } }

func WithMaxRetries(n int) Option { return func(c *Client) { c.maxRetries = n } }

func New(tokens TokenSource, opts ...Option) *Client {
	c := &Client{
		http:       &http.Client{Timeout: 60 * time.Second},
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

// Do выполняет запрос к Graph. path — относительный ("/users") или абсолютный URL
// (например, @odata.nextLink). Ответ декодируется в out (если out != nil и есть тело).
func (c *Client) Do(ctx context.Context, method, path string, params url.Values, body any, out any) error {
	raw, err := c.doRaw(ctx, method, path, params, body)
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

func (c *Client) doRaw(ctx context.Context, method, path string, params url.Values, body any) ([]byte, error) {
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
			return nil, fmt.Errorf("graph: encode request: %w", err)
		}
		payload = b
	}

	method = strings.ToUpper(method)

	for attempt := 0; ; attempt++ {
		raw, retryAfter, err := c.once(ctx, method, u, payload)
		if err == nil {
			return raw, nil
		}

		if attempt >= c.maxRetries || !retryable(method, err) {
			return nil, err
		}

		delay := retryAfter
		if delay <= 0 {
			delay = time.Second << attempt // 1s, 2s, 4s, 8s
		}
		if serr := c.sleep(ctx, delay); serr != nil {
			return nil, serr
		}
	}
}

// retryable: 429 (throttling) ретраим для любых методов — Graph запрос не обработал;
// 503/504 — только идемпотентные методы.
func retryable(method string, err error) bool {
	ge, ok := err.(*GraphError)
	if !ok {
		return false
	}
	switch ge.StatusCode {
	case 429:
		return true
	case 503, 504:
		return method == http.MethodGet || method == http.MethodPut || method == http.MethodDelete
	}
	return false
}

func (c *Client) once(ctx context.Context, method, u string, payload []byte) ([]byte, time.Duration, error) {
	var bodyReader io.Reader
	if payload != nil {
		bodyReader = bytes.NewReader(payload)
	}
	req, err := http.NewRequestWithContext(ctx, method, u, bodyReader)
	if err != nil {
		return nil, 0, err
	}

	token, err := c.tokens.Token(ctx)
	if err != nil {
		return nil, 0, fmt.Errorf("graph: acquire token: %w", err)
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "application/json")
	if payload != nil {
		req.Header.Set("Content-Type", "application/json")
	}

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, 0, err
	}
	defer resp.Body.Close()

	raw, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, 0, err
	}

	if resp.StatusCode >= 200 && resp.StatusCode < 300 {
		if resp.StatusCode == http.StatusNoContent {
			return nil, 0, nil
		}
		return raw, 0, nil
	}

	return nil, parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, raw)
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
		// не-JSON тело: не тащим потенциально чувствительный дамп целиком
		const maxSnippet = 200
		s := string(raw)
		if len(s) > maxSnippet {
			s = s[:maxSnippet] + "..."
		}
		ge.Message = s
	}
	return ge
}

// Удобные шорткаты.

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
