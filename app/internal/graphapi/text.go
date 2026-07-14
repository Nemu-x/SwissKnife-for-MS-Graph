package graphapi

import (
	"context"
	"io"
	"net/http"
	"net/url"
	"strings"
)

// GetText fetches a resource as plain text (e.g. Graph usage reports, which
// return CSV). Redirects to the download URL are followed by the http client.
func (c *Client) GetText(ctx context.Context, path string, params url.Values) (string, error) {
	u := path
	if !strings.HasPrefix(path, "http") {
		u = c.baseURL + "/" + strings.TrimLeft(path, "/")
	}
	if len(params) > 0 {
		sep := "?"
		if strings.Contains(u, "?") {
			sep = "&"
		}
		u += sep + params.Encode()
	}

	token, err := c.tokens.Token(ctx)
	if err != nil {
		return "", err
	}
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("Authorization", "Bearer "+token)

	resp, err := c.http.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	raw, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", parseGraphError(resp, raw)
	}
	return string(raw), nil
}
