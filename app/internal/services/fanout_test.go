package services

import (
	"context"
	"errors"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/secrets"
)

func TestRunAcrossTagsRowsAndNamesFailedTenants(t *testing.T) {
	guests := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"value":[{"id":"g1","displayName":"Guest One","userPrincipalName":"one_x.com#EXT#@a.onmicrosoft.com","mail":"one@x.com"}]}`))
	}))
	t.Cleanup(guests.Close)
	prev := profileEnv
	t.Cleanup(func() { profileEnv = prev })
	profileEnv = func(ctx context.Context, _ *secrets.Store, id string) (string, engine.Env, error) {
		if id == "broken" {
			return "Contoso", engine.Env{}, errors.New("AADSTS7000215: invalid client secret")
		}
		return "Fabrikam " + id, engine.Env{Ctx: ctx, Graph: graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(guests.URL)), AppOnly: true}, nil
	}
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	a := NewActionsService(sess, secrets.NewStoreAt(t.TempDir()))

	res, err := a.RunAcross("report.guests", nil, []string{"a", "broken", "b"})
	if err != nil {
		t.Fatal(err)
	}
	if res.Columns[0] != "tenant" || len(res.Rows) != 2 || res.Rows[0]["tenant"] != "Fabrikam a" || res.Rows[1]["tenant"] != "Fabrikam b" {
		t.Fatalf("result %+v", res)
	}
	if res.Note == nil || !strings.Contains(res.Note.Params["list"], "Contoso") {
		t.Fatalf("the failed tenant must be named: %+v", res.Note)
	}
	if _, err := a.RunAcross("group.membership", nil, []string{"a"}); err == nil {
		t.Fatal("a write action cannot fan out")
	}
}
