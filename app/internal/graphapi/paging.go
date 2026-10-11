package graphapi

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
)

// Page is a single page of a Graph list response.
type Page struct {
	Items    []json.RawMessage
	NextLink string
}

type listEnvelope struct {
	Value    []json.RawMessage `json:"value"`
	NextLink string            `json:"@odata.nextLink"`
}

// ListPage returns a single page. path is a relative path or a nextLink.
func (c *Client) ListPage(ctx context.Context, path string, params url.Values) (*Page, error) {
	var env listEnvelope
	if err := c.Do(ctx, http.MethodGet, path, params, nil, &env); err != nil {
		return nil, err
	}
	return &Page{Items: env.Value, NextLink: env.NextLink}, nil
}

// ListAll collects every page following @odata.nextLink.
// maxItems > 0 is a safety cap (ADR-003); 0 means unlimited.
func (c *Client) ListAll(ctx context.Context, path string, params url.Values, maxItems int) ([]json.RawMessage, error) {
	items, _, err := c.ListAllCapped(ctx, path, params, maxItems)
	return items, err
}

// ListAllCapped is ListAll that also reports whether maxItems cut the
// collection short (items or pages remained beyond the cap), so a caller can
// tell "the tenant has exactly N" from "we stopped at N".
func (c *Client) ListAllCapped(ctx context.Context, path string, params url.Values, maxItems int) ([]json.RawMessage, bool, error) {
	var items []json.RawMessage
	next := path
	p := params
	for next != "" {
		page, err := c.ListPage(ctx, next, p)
		if err != nil {
			return nil, false, err
		}
		items = append(items, page.Items...)
		if maxItems > 0 && len(items) >= maxItems {
			truncated := len(items) > maxItems || page.NextLink != ""
			return items[:maxItems], truncated, nil
		}
		// nextLink already carries the query, so do not pass params again
		next, p = page.NextLink, nil
	}
	return items, false, nil
}

// ListAllInto is ListAll that decodes each item into []T.
func ListAllInto[T any](ctx context.Context, c *Client, path string, params url.Values, maxItems int) ([]T, error) {
	raw, err := c.ListAll(ctx, path, params, maxItems)
	if err != nil {
		return nil, err
	}
	out := make([]T, 0, len(raw))
	for i, r := range raw {
		var v T
		if err := json.Unmarshal(r, &v); err != nil {
			return nil, fmt.Errorf("graph: decode list item %d: %w", i, err)
		}
		out = append(out, v)
	}
	return out, nil
}
