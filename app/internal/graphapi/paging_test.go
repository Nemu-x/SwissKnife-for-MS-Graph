package graphapi

import (
	"context"
	"fmt"
	"net/http"
	"testing"
)

func TestListAllCappedReportsTruncation(t *testing.T) {
	// Two pages of two items each.
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Query().Get("page") == "2" {
			fmt.Fprint(w, `{"value":[{"id":"3"},{"id":"4"}]}`)
			return
		}
		fmt.Fprintf(w, `{"value":[{"id":"1"},{"id":"2"}],"@odata.nextLink":"http://%s/items?page=2"}`, r.Host)
	}))
	ctx := context.Background()
	cases := []struct {
		max       int
		wantN     int
		truncated bool
	}{
		{0, 4, false}, // unlimited
		{4, 4, false}, // cap equals the collection: nothing was cut
		{5, 4, false},
		{3, 3, true}, // cut inside the second page
		{2, 2, true}, // cut at a page boundary with a nextLink left
	}
	for _, tc := range cases {
		items, truncated, err := c.ListAllCapped(ctx, "/items", nil, tc.max)
		if err != nil {
			t.Fatal(err)
		}
		if len(items) != tc.wantN || truncated != tc.truncated {
			t.Errorf("max=%d: got %d items truncated=%v, want %d truncated=%v", tc.max, len(items), truncated, tc.wantN, tc.truncated)
		}
	}
	// ListAll keeps its old contract.
	if items, err := c.ListAll(ctx, "/items", nil, 3); err != nil || len(items) != 3 {
		t.Errorf("ListAll = %d, %v", len(items), err)
	}
}
