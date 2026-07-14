package graphapi

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
)

// Page — одна страница списочного ответа Graph.
type Page struct {
	Items    []json.RawMessage
	NextLink string
}

type listEnvelope struct {
	Value    []json.RawMessage `json:"value"`
	NextLink string            `json:"@odata.nextLink"`
}

// ListPage возвращает одну страницу. path — относительный путь или nextLink.
func (c *Client) ListPage(ctx context.Context, path string, params url.Values) (*Page, error) {
	var env listEnvelope
	if err := c.Do(ctx, http.MethodGet, path, params, nil, &env); err != nil {
		return nil, err
	}
	return &Page{Items: env.Value, NextLink: env.NextLink}, nil
}

// ListAll собирает все страницы по @odata.nextLink.
// maxItems > 0 — защитный лимит (safety, ADR-003); 0 — без лимита.
func (c *Client) ListAll(ctx context.Context, path string, params url.Values, maxItems int) ([]json.RawMessage, error) {
	var items []json.RawMessage
	next := path
	p := params
	for next != "" {
		page, err := c.ListPage(ctx, next, p)
		if err != nil {
			return nil, err
		}
		items = append(items, page.Items...)
		if maxItems > 0 && len(items) >= maxItems {
			return items[:maxItems], nil
		}
		// nextLink уже содержит query — параметры повторно не передаём
		next, p = page.NextLink, nil
	}
	return items, nil
}

// ListAllInto — ListAll с декодированием каждого элемента в []T.
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
