package graphapi

import (
	"context"
	"io"
	"net/http"
	"strconv"
	"strings"
)

// Count returns the collection count via /$count. Graph requires the
// ConsistencyLevel: eventual header for advanced queries on /users and /groups.
func (c *Client) Count(ctx context.Context, collection string) (int, error) {
	u := c.baseURL + "/" + strings.Trim(collection, "/") + "/$count"

	token, err := c.tokens.Token(ctx)
	if err != nil {
		return 0, err
	}
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return 0, err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "text/plain")
	req.Header.Set("ConsistencyLevel", "eventual")

	resp, err := c.http.Do(req)
	if err != nil {
		return 0, err
	}
	defer resp.Body.Close()
	raw, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return 0, parseGraphError(resp, raw)
	}
	n, err := strconv.Atoi(strings.TrimSpace(string(raw)))
	if err != nil {
		return 0, nil
	}
	return n, nil
}
