package session

import (
	"context"
	"sync"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
)

// tenantTokens tags a token broker with the tenant it belongs to.
type tenantTokens string

func (tenantTokens) TokenFor(context.Context, string) (string, error) { return "t", nil }

func conn(tenant string) Conn {
	return Conn{
		Client: graphapi.New(graphapi.StaticToken(tenant)), ProfileName: "p-" + tenant,
		Tokens: tenantTokens(tenant), TenantID: tenant, AppOnly: tenant == "a",
		ProfileID: "id-" + tenant, DelegatedOrg: "org-" + tenant,
		Policy: Policy{MaxDanger: "read", AllowedGroups: []string{"g-" + tenant}},
	}
}

func TestConnectOverAConnectionRunsDisconnectHooks(t *testing.T) {
	s := New(auditlog.New(t.TempDir()))
	var ended []*graphapi.Client
	s.OnDisconnect(func(prev *graphapi.Client) {
		// The hooks run before the new connection is visible.
		if s.Connected() {
			t.Error("hooks must see the old connection gone, not the new one in place")
		}
		ended = append(ended, prev)
	})
	a, b := conn("a"), conn("b")
	s.Connect(a)
	if len(ended) != 0 {
		t.Fatal("a first connect ends nothing")
	}
	s.Connect(b)
	if len(ended) != 1 || ended[0] != a.Client {
		t.Fatalf("switching must end the previous connection through its hooks: %v", ended)
	}
	if got := s.Current(); got.Client != b.Client || got.TenantID != "b" || got.ProfileID != "id-b" || got.Policy.AllowedGroups[0] != "g-b" {
		t.Fatalf("current = %+v", got)
	}
}

// Readers never see a part of one connection next to a part of another.
func TestCurrentNeverSeesAHalfSwappedConnection(t *testing.T) {
	s := New(auditlog.New(t.TempDir()))
	conns := map[string]Conn{"a": conn("a"), "b": conn("b")}
	s.Connect(conns["a"])
	stop := make(chan struct{})
	var wg sync.WaitGroup
	wg.Add(1)
	go func() {
		defer wg.Done()
		for i := 0; ; i++ {
			select {
			case <-stop:
				return
			default:
			}
			if i%2 == 0 {
				s.Connect(conns["b"])
			} else {
				s.Connect(conns["a"])
			}
		}
	}()
	for i := 0; i < 20000; i++ {
		c := s.Current()
		if c.Client == nil {
			continue // between the end of one connection and the next
		}
		tenant := string(c.Tokens.(tenantTokens))
		want := conns[tenant]
		if c.Client != want.Client || c.TenantID != tenant || c.AppOnly != want.AppOnly ||
			c.ProfileID != want.ProfileID || c.DelegatedOrg != want.DelegatedOrg || c.ProfileName != want.ProfileName ||
			c.Policy.AllowedGroups[0] != "g-"+tenant {
			close(stop)
			wg.Wait()
			t.Fatalf("mixed connection: %+v", c)
		}
	}
	close(stop)
	wg.Wait()
}
