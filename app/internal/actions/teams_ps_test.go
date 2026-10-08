package actions

import (
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// teamsFake answers Teams cmdlets from a table and records every call.
type teamsFake struct {
	answers map[string]string
	calls   []exoCall
}

func (f *teamsFake) Invoke(_ engine.Env, family, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	if family != pwsh.FamilyTeams {
		panic("unexpected family " + family)
	}
	f.calls = append(f.calls, exoCall{cmdlet, params})
	var out []json.RawMessage
	if a, ok := f.answers[cmdlet]; ok {
		_ = json.Unmarshal([]byte(a), &out)
	}
	return out, nil
}

func (f *teamsFake) writes() []exoCall {
	var out []exoCall
	for _, c := range f.calls {
		if !strings.HasPrefix(c.Cmdlet, "Get-") {
			out = append(out, c)
		}
	}
	return out
}

func teamsHarness(t *testing.T, fake *teamsFake) *engine.Engine {
	t.Helper()
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == "/groups/g1" {
			w.Write([]byte(`{"id":"g1","displayName":"Sales"}`))
			return
		}
		graphUsers(w, r)
	}))
	t.Cleanup(srv.Close)
	s := session.New(auditlog.New(t.TempDir()))
	s.SetClient(graphapi.New(graphapi.StaticToken("t"), graphapi.WithBaseURL(srv.URL)), "test")
	e := engine.New(s, engine.GraphProvider{}, stubProvider{backend: pwsh.BackendTeamsPS})
	e.PS = fake
	e.Register(Builtin()...)
	return e
}

func TestTeamsUserPolicyGrantAndReset(t *testing.T) {
	fake := &teamsFake{answers: map[string]string{
		"Get-CsTeamsMeetingPolicy": `[{"Identity":"Tag:NoRecording"}]`,
		"Get-CsOnlineUser":         `[{"TeamsMeetingPolicy":null}]`,
	}}
	e := teamsHarness(t, fake)
	ch := planApply(t, e, "teams.userPolicy", engine.Inputs{"user": "bob@contoso.com", "policy": "NoRecording"})
	if ch.Before != "Global" || ch.After != "NoRecording" || ch.Field != "teamsPolicy.meeting" {
		t.Fatalf("change %+v", ch)
	}
	if w := fake.writes(); len(w) != 1 || w[0].Cmdlet != "Grant-CsTeamsMeetingPolicy" || w[0].Params["PolicyName"] != "NoRecording" {
		t.Fatalf("writes %+v", w)
	}

	// Back to the org default: PolicyName must be null, not the word "Global".
	fake = &teamsFake{answers: map[string]string{"Get-CsOnlineUser": `[{"TeamsMeetingPolicy":{"Name":"Tag:NoRecording"}}]`}}
	planApply(t, teamsHarness(t, fake), "teams.userPolicy", engine.Inputs{"user": "bob@contoso.com", "policy": "Global"})
	w := fake.writes()
	if v, ok := w[0].Params["PolicyName"]; !ok || v != nil {
		t.Fatalf("reset must send a null policy: %+v", w)
	}
}

func TestTeamsUnknownPolicyIsRefused(t *testing.T) {
	fake := &teamsFake{answers: map[string]string{"Get-CsOnlineUser": `[{}]`}}
	if _, err := teamsHarness(t, fake).Plan("teams.userPolicy", engine.Inputs{"user": "bob@contoso.com", "policy": "Typo"}); err == nil {
		t.Fatal("a policy that does not exist must be refused at preview")
	}
}

func TestTeamsGroupPolicyTakesTheNextRank(t *testing.T) {
	fake := &teamsFake{answers: map[string]string{
		"Get-CsTeamsMessagingPolicy":  `[{"Identity":"Tag:Strict"}]`,
		"Get-CsGroupPolicyAssignment": `[{"GroupId":"other","PolicyName":"Loose","Rank":1},{"GroupId":"x","PolicyName":"Loose","Rank":4}]`,
	}}
	ch := planApply(t, teamsHarness(t, fake), "teams.groupPolicy", engine.Inputs{"group": "g1", "policyType": "messaging", "policy": "Strict"})
	if ch.Op != "add" || ch.Target != "Sales" {
		t.Fatalf("change %+v", ch)
	}
	w := fake.writes()
	if len(w) != 1 || w[0].Cmdlet != "New-CsGroupPolicyAssignment" || w[0].Params["Rank"] != "5" || w[0].Params["PolicyType"] != "TeamsMessagingPolicy" {
		t.Fatalf("writes %+v", w)
	}
}

func TestTeamsEffectivePoliciesRead(t *testing.T) {
	fake := &teamsFake{answers: map[string]string{"Get-CsOnlineUser": `[{"TeamsMeetingPolicy":"Tag:NoRecording","TeamsCallingPolicy":null}]`}}
	res, err := teamsHarness(t, fake).Run(t.Context(), "teams.effectivePolicies", engine.Inputs{"user": "bob@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != len(policyTypes) || res.Rows[0]["policy"] != "NoRecording" || res.Rows[2]["policy"] != "Global" {
		t.Fatalf("rows %+v", res.Rows)
	}
	if len(fake.writes()) != 0 {
		t.Fatal("a read must not write")
	}
}

func TestEveryTeamsCmdletIsAllowListed(t *testing.T) {
	allowed := map[string]bool{}
	for _, c := range PowerShellCmdlets()[pwsh.FamilyTeams] {
		allowed[c] = true
	}
	answers := map[string]string{"Get-CsOnlineUser": `[{}]`, "Get-CsGroupPolicyAssignment": `[{"GroupId":"g1","PolicyName":"A","Rank":1}]`}
	for _, p := range teamsPolicies {
		answers[p.get] = `[{"Identity":"Tag:A"}]`
	}
	fake := &teamsFake{answers: answers}
	e := teamsHarness(t, fake)
	for _, typ := range policyTypes {
		planApply(t, e, "teams.userPolicy", engine.Inputs{"user": "bob@contoso.com", "policyType": typ, "policy": "A"})
		planApply(t, e, "teams.groupPolicy", engine.Inputs{"group": "g1", "policyType": typ, "policy": "Global"})
		if _, err := e.Run(t.Context(), "teams.policies", engine.Inputs{"policyType": typ}); err != nil {
			t.Fatal(err)
		}
	}
	for _, c := range fake.calls {
		if !allowed[c.Cmdlet] {
			t.Errorf("%s is called but not allow-listed", c.Cmdlet)
		}
	}
}
