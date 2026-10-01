package graphapi

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net"
	"net/http"
	"net/url"
	"time"
)

// PostPreauthorized POSTs a JSON body to an absolute, pre-authenticated URL and
// decodes the JSON response into out (when non-nil). Graph hands such URLs out
// for session-style uploads — the mailbox import URL carries its own token in
// the query string and the docs explicitly say not to send an Authorization
// header with it — so unlike Do no Bearer token is attached. Throttling (429)
// and transport failures retry exactly like ordinary Graph calls, and error
// bodies parse into *GraphError.
func (c *Client) PostPreauthorized(ctx context.Context, u string, body, out any) error {
	parsed, perr := url.Parse(u)
	if perr != nil || !parsed.IsAbs() || (parsed.Scheme != "https" && parsed.Scheme != "http") {
		return errors.New("graph: pre-authorized URL must be absolute")
	}
	// The token rides in the query string: plain HTTP would put it on the wire
	// in clear. Loopback stays allowed for the httptest-based tests.
	if parsed.Scheme == "http" && !isLoopback(parsed.Hostname()) {
		return errors.New("graph: pre-authorized URL must use https")
	}
	var payload []byte
	if body != nil {
		b, err := json.Marshal(body)
		if err != nil {
			return fmt.Errorf("graph: encode request: %w", err)
		}
		payload = b
	}
	var raw []byte
	err := c.withRetry(ctx, http.MethodPost, func() (time.Duration, error) {
		var rd io.Reader
		if payload != nil {
			rd = bytes.NewReader(payload)
		}
		req, err := http.NewRequestWithContext(ctx, http.MethodPost, u, rd)
		if err != nil {
			return 0, err
		}
		req.Header.Set("Accept", "application/json")
		if payload != nil {
			req.Header.Set("Content-Type", "application/json")
		}
		// Never follow redirects: the https/loopback rule above was checked for
		// this URL only, and Go would re-send the token (Referer, or the whole
		// body on 307/308) to wherever a 3xx points. A 3xx comes back as an error.
		noRedirect := *c.http
		noRedirect.CheckRedirect = func(*http.Request, []*http.Request) error { return http.ErrUseLastResponse }
		resp, err := noRedirect.Do(req)
		if err != nil {
			// Transport errors quote the request URL, token included; strip the
			// query so the text can travel to reports and the UI (ADR-002).
			var ue *url.Error
			if errors.As(err, &ue) {
				ue.URL = parsed.Scheme + "://" + parsed.Host + parsed.Path
			}
			return 0, err
		}
		defer resp.Body.Close()
		b, err := io.ReadAll(resp.Body)
		if err != nil {
			return 0, err
		}
		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			return parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, b)
		}
		raw = b
		return 0, nil
	})
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

// isLoopback reports whether host is localhost or a loopback IP.
func isLoopback(host string) bool {
	if host == "localhost" {
		return true
	}
	ip := net.ParseIP(host)
	return ip != nil && ip.IsLoopback()
}
