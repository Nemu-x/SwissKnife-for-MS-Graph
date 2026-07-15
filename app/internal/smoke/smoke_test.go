//go:build smoke

// Package smoke runs READ-ONLY checks against a real Microsoft 365 tenant.
// It exercises what the httptest fakes cannot: real auth, real Graph behavior,
// real permissions. Excluded from normal `go test` by the build tag.
//
// Run: go test -tags=smoke ./internal/smoke/ -v
// Env: SMOKE_TENANT_ID, SMOKE_CLIENT_ID, SMOKE_CLIENT_SECRET (app-only creds
// of a TEST tenant — never point this at production).
package smoke

import (
	"context"
	"os"
	"testing"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/graphapi"
)

func client(t *testing.T) *graphapi.Client {
	t.Helper()
	tenant := os.Getenv("SMOKE_TENANT_ID")
	id := os.Getenv("SMOKE_CLIENT_ID")
	secret := os.Getenv("SMOKE_CLIENT_SECRET")
	if tenant == "" || id == "" || secret == "" {
		t.Skip("SMOKE_* env not set — skipping real-tenant smoke")
	}
	tp, err := auth.NewClientSecret(tenant, id, secret)
	if err != nil {
		t.Fatalf("credential: %v", err)
	}
	return graphapi.New(tp)
}

func ctx(t *testing.T) context.Context {
	c, cancel := context.WithTimeout(context.Background(), 60*time.Second)
	t.Cleanup(cancel)
	return c
}

func TestOrganizationReadable(t *testing.T) {
	c := client(t)
	var out struct {
		Value []struct {
			DisplayName string `json:"displayName"`
		} `json:"value"`
	}
	if err := c.Get(ctx(t), "/organization", nil, &out); err != nil {
		t.Fatalf("GET /organization: %v", err)
	}
	if len(out.Value) == 0 || out.Value[0].DisplayName == "" {
		t.Fatal("organization came back empty")
	}
	t.Logf("tenant: %s", out.Value[0].DisplayName)
}

func TestUsersListAndPaging(t *testing.T) {
	c := client(t)
	users, err := c.ListAll(ctx(t), "/users", nil, 5)
	if err != nil {
		t.Fatalf("list users: %v", err)
	}
	if len(users) == 0 {
		t.Fatal("no users returned — check User.Read.All + admin consent")
	}
}

func TestSubscribedSkus(t *testing.T) {
	c := client(t)
	var out struct {
		Value []struct {
			SkuPartNumber string `json:"skuPartNumber"`
		} `json:"value"`
	}
	if err := c.Get(ctx(t), "/subscribedSkus", nil, &out); err != nil {
		t.Fatalf("GET /subscribedSkus: %v", err)
	}
}

func TestServicePrincipalsReadable(t *testing.T) {
	c := client(t)
	if _, err := c.ListAll(ctx(t), "/servicePrincipals", nil, 5); err != nil {
		t.Fatalf("list servicePrincipals (Security page): %v", err)
	}
}
