package actions

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

type call struct{ method, path, body string }

// harness wires an engine with the built-in actions to a fake Graph.
func harness(t *testing.T, h func(w http.ResponseWriter, r *http.Request)) (*engine.Engine, *[]call) {
	t.Helper()
	var calls []call
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		calls = append(calls, call{r.Method, r.URL.Path, string(b)})
		h(w, r)
	}))
	t.Cleanup(srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{})
	e.Register(Builtin()...)
	return e, &calls
}

func writes(calls []call) []call {
	var out []call
	for _, c := range calls {
		if c.method != "GET" {
			out = append(out, c)
		}
	}
	return out
}

func TestSignInAlreadyBlockedIsNoOp(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","accountEnabled":false}`))
	})
	p, err := e.Plan("user.signIn", engine.Inputs{"user": "ann@contoso.com", "state": "blocked"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Op != "none" {
		t.Fatalf("change = %+v, want none", p.Changes[0])
	}
	r, err := e.Apply(p.ID, "")
	if err != nil || r.Skipped != 1 || len(writes(*calls)) != 0 {
		t.Fatalf("result = %+v err = %v writes = %v", r, err, writes(*calls))
	}
}

func TestSignInUnblockPatchesById(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","accountEnabled":false}`))
	})
	p, _ := e.Plan("user.signIn", engine.Inputs{"user": "ann@contoso.com", "state": "allowed"})
	if c := p.Changes[0]; c.Op != "set" || c.Before != "blocked" || c.After != "allowed" {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writes(*calls)
	if len(w) != 1 || w[0].method != "PATCH" || w[0].path != "/users/u1" || w[0].body != `{"accountEnabled":true}` {
		t.Fatalf("writes = %+v", w)
	}
}

func TestGroupMembershipPlans(t *testing.T) {
	handler := func(member bool) func(w http.ResponseWriter, r *http.Request) {
		return func(w http.ResponseWriter, r *http.Request) {
			switch r.URL.Path {
			case "/users/ann@contoso.com":
				w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com"}`))
			case "/groups/g1":
				w.Write([]byte(`{"id":"g1","displayName":"Sales"}`))
			case "/users/u1/memberOf":
				if member {
					w.Write([]byte(`{"value":[{"id":"g9"},{"id":"g1"}]}`))
				} else {
					w.Write([]byte(`{"value":[{"id":"g9"}]}`))
				}
			default:
				w.WriteHeader(http.StatusNoContent)
			}
		}
	}
	in := engine.Inputs{"user": "ann@contoso.com", "group": "g1"}

	e, _ := harness(t, handler(true))
	p, err := e.Plan("group.membership", in)
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "none" || c.Field != "group.member" {
		t.Fatalf("already a member: change = %+v", c)
	}

	e, calls := harness(t, handler(false))
	p, _ = e.Plan("group.membership", in)
	if c := p.Changes[0]; c.Op != "add" || c.After != "Sales" {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writes(*calls)
	if len(w) != 1 {
		t.Fatalf("writes = %+v", w)
	}
	var body map[string]string
	_ = json.Unmarshal([]byte(w[0].body), &body)
	if w[0].path != "/groups/g1/members/$ref" || body["@odata.id"] != "https://graph.microsoft.com/v1.0/directoryObjects/u1" {
		t.Fatalf("writes = %+v", w)
	}

	e, calls = harness(t, handler(true))
	p, _ = e.Plan("group.membership", engine.Inputs{"user": "ann@contoso.com", "group": "g1", "op": "remove"})
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	if w := writes(*calls); len(w) != 1 || w[0].method != "DELETE" || w[0].path != "/groups/g1/members/u1/$ref" {
		t.Fatalf("remove writes = %+v", w)
	}
}

func TestLicensePlanNamesSku(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users/ann@contoso.com":
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","licenseAssignmentStates":[]}`))
		case "/subscribedSkus":
			w.Write([]byte(`{"value":[{"skuId":"s1","skuPartNumber":"ENTERPRISEPACK"}]}`))
		default:
			w.Write([]byte(`{}`))
		}
	})
	p, err := e.Plan("license.assign", engine.Inputs{"user": "ann@contoso.com", "sku": "s1"})
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "add" || c.After != "ENTERPRISEPACK" {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writes(*calls)
	if len(w) != 1 || w[0].path != "/users/u1/assignLicense" || w[0].body != `{"addLicenses":[{"skuId":"s1"}],"removeLicenses":[]}` {
		t.Fatalf("writes = %+v", w)
	}

	// Removing a license the user does not hold is a no-op.
	p, _ = e.Plan("license.assign", engine.Inputs{"user": "ann@contoso.com", "sku": "s1", "op": "remove"})
	if p.Changes[0].Op != "none" {
		t.Fatalf("remove absent: %+v", p.Changes[0])
	}
}

func TestRevokeSessionsNeedsConfirm(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com"}`))
	})
	p, _ := e.Plan("user.revokeSessions", engine.Inputs{"user": "ann@contoso.com"})
	if _, err := e.Apply(p.ID, "nope"); err == nil {
		t.Fatal("must refuse without confirm")
	}
	p, _ = e.Plan("user.revokeSessions", engine.Inputs{"user": "ann@contoso.com"})
	if _, err := e.Apply(p.ID, "ann@contoso.com"); err != nil {
		t.Fatal(err)
	}
	if w := writes(*calls); len(w) != 1 || w[0].path != "/users/u1/revokeSignInSessions" {
		t.Fatalf("writes = %+v", w)
	}
}

func TestLicenseInheritedFromGroupIsNotRemovable(t *testing.T) {
	e, _ := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users/ann@contoso.com":
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","licenseAssignmentStates":[{"skuId":"S1","assignedByGroup":"g1"}]}`))
		default:
			w.Write([]byte(`{"value":[{"skuId":"s1","skuPartNumber":"ENTERPRISEPACK"}]}`))
		}
	})
	p, err := e.Plan("license.assign", engine.Inputs{"user": "ann@contoso.com", "sku": "s1", "op": "remove"})
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "none" || c.Note != "inheritedLicense" || c.Before != "ENTERPRISEPACK" {
		t.Fatalf("change = %+v", c)
	}
}

func TestManagerPlanAndApply(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users/ann@contoso.com":
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com"}`))
		case "/users/boss@contoso.com":
			w.Write([]byte(`{"id":"m1","userPrincipalName":"boss@contoso.com"}`))
		case "/users/u1/manager":
			w.WriteHeader(http.StatusNotFound)
			w.Write([]byte(`{"error":{"code":"Request_ResourceNotFound","message":"no manager"}}`))
		default:
			w.WriteHeader(http.StatusNoContent)
		}
	})
	p, err := e.Plan("user.manager", engine.Inputs{"user": "ann@contoso.com", "manager": "boss@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if c := p.Changes[0]; c.Op != "set" || c.Before != "" || c.After != "boss@contoso.com" {
		t.Fatalf("change = %+v", c)
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	w := writes(*calls)
	if len(w) != 1 || w[0].method != "PUT" || w[0].path != "/users/u1/manager/$ref" {
		t.Fatalf("writes = %+v", w)
	}
}

func TestUsageLocationValidatesAndUppercases(t *testing.T) {
	e, calls := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "GET" {
			w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","usageLocation":"US"}`))
			return
		}
		w.WriteHeader(http.StatusNoContent)
	})
	if _, err := e.Plan("user.usageLocation", engine.Inputs{"user": "ann@contoso.com", "country": "Germany"}); err == nil {
		t.Fatal("a non-ISO value must be rejected")
	}
	p, _ := e.Plan("user.usageLocation", engine.Inputs{"user": "ann@contoso.com", "country": "us"})
	if p.Changes[0].Op != "none" {
		t.Fatalf("same country: %+v", p.Changes[0])
	}
	p, _ = e.Plan("user.usageLocation", engine.Inputs{"user": "ann@contoso.com", "country": "de"})
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	if w := writes(*calls); len(w) != 1 || w[0].body != `{"usageLocation":"DE"}` {
		t.Fatalf("writes = %+v", w)
	}
}
