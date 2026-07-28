package services

import (
	"encoding/json"
	"net/http"
	"strings"
	"testing"
)

// mirrorFake serves a tenant where the source user has more access than the
// target: two groups (one of them a distribution list Graph cannot touch), an
// admin role, two teams, a private channel and a license.
func mirrorFake(t *testing.T, calls *[]string) http.HandlerFunc {
	t.Helper()
	return func(w http.ResponseWriter, r *http.Request) {
		if r.Method != "GET" {
			body, _ := json.Marshal(map[string]any{})
			*calls = append(*calls, r.Method+" "+r.URL.Path)
			w.Write(body)
			return
		}
		switch r.URL.Path {
		case "/users/src@contoso.com":
			w.Write([]byte(`{"id":"src-id","userPrincipalName":"src@contoso.com"}`))
		case "/users/dst@contoso.com":
			w.Write([]byte(`{"id":"dst-id","userPrincipalName":"dst@contoso.com"}`))

		case "/users/src-id/memberOf":
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.group","id":"g1","displayName":"Sales","groupTypes":["Unified"],"mailEnabled":true},
				{"@odata.type":"#microsoft.graph.group","id":"g2","displayName":"All staff DL","groupTypes":[],"mailEnabled":true},
				{"@odata.type":"#microsoft.graph.group","id":"g3","displayName":"Dyn","groupTypes":["DynamicMembership"],"securityEnabled":true},
				{"@odata.type":"#microsoft.graph.directoryRole","id":"r1","displayName":"Helpdesk Administrator"}]}`))
		case "/users/dst-id/memberOf":
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.group","id":"g1","displayName":"Sales","groupTypes":["Unified"],"mailEnabled":true},
				{"@odata.type":"#microsoft.graph.group","id":"g9","displayName":"Target only","securityEnabled":true}]}`))

		case "/users/src-id/joinedTeams":
			w.Write([]byte(`{"value":[{"id":"t1","displayName":"HelpCenter"},{"id":"t2","displayName":"Board"}]}`))
		case "/users/dst-id/joinedTeams":
			w.Write([]byte(`{"value":[{"id":"t1","displayName":"HelpCenter"}]}`))

		case "/teams/t1/allChannels":
			w.Write([]byte(`{"value":[
				{"id":"c0","displayName":"General","membershipType":"standard"},
				{"id":"c1","displayName":"Escalations","membershipType":"private"}]}`))
		case "/teams/t2/allChannels":
			w.Write([]byte(`{"value":[{"id":"c2","displayName":"Secrets","membershipType":"private"}]}`))
		case "/teams/t1/channels/c1/members":
			w.Write([]byte(`{"value":[{"userId":"src-id","email":"src@contoso.com"}]}`))
		case "/teams/t2/channels/c2/members":
			w.Write([]byte(`{"value":[{"userId":"src-id","email":"src@contoso.com"}]}`))

		case "/users/src-id/licenseDetails":
			w.Write([]byte(`{"value":[{"skuId":"sku-e3","skuPartNumber":"ENTERPRISEPACK"}]}`))
		case "/users/dst-id/licenseDetails":
			w.Write([]byte(`{"value":[]}`))
		default:
			w.Write([]byte(`{"value":[]}`))
		}
	}
}

func find(rows []AccessRow, kind, id string) *AccessRow {
	for i := range rows {
		if rows[i].Kind == kind && rows[i].ID == id {
			return &rows[i]
		}
	}
	return nil
}

func TestMirrorCompareClassifiesEverySideOfTheDiff(t *testing.T) {
	var calls []string
	sess := harness(t, mirrorFake(t, &calls))
	rows, err := NewMirrorService(sess).Compare("src@contoso.com", "dst@contoso.com")
	if err != nil {
		t.Fatal(err)
	}

	// Shared group: present on both sides, so nothing to copy.
	if g := find(rows, KindGroup, "g1"); g == nil || g.Status != "both" || g.Copyable {
		t.Errorf("g1: want status=both copyable=false, got %+v", g)
	}
	// Exchange-managed distribution list: missing, but Graph cannot add to it.
	if g := find(rows, KindGroup, "g2"); g == nil || g.Status != "missing" || g.Copyable || g.ReasonKey != "reasons.exchangeGroup" {
		t.Errorf("g2: want missing + not copyable + exchange reason, got %+v", g)
	}
	// Dynamic group: membership comes from a rule.
	if g := find(rows, KindGroup, "g3"); g == nil || g.Copyable || g.ReasonKey != "reasons.dynamicGroup" {
		t.Errorf("g3: want not copyable + dynamic reason, got %+v", g)
	}
	if r := find(rows, KindRole, "r1"); r == nil || r.Status != "missing" || !r.Copyable {
		t.Errorf("r1: want missing + copyable, got %+v", r)
	}
	if tm := find(rows, KindTeam, "t2"); tm == nil || tm.Status != "missing" {
		t.Errorf("t2: want missing, got %+v", tm)
	}
	// Standard channels have no membership of their own and must not be listed.
	if c := find(rows, KindChannel, "c0"); c != nil {
		t.Errorf("standard channel must not appear in the diff: %+v", c)
	}
	// Private channel of a shared team, held by the source only.
	if c := find(rows, KindChannel, "c1"); c == nil || c.Status != "missing" || c.TeamID != "t1" {
		t.Errorf("c1: want missing with teamId=t1, got %+v", c)
	}
	if l := find(rows, KindLicense, "sku-e3"); l == nil || l.Status != "missing" {
		t.Errorf("license: want missing, got %+v", l)
	}
	// What only the target has is reported, never copied.
	if g := find(rows, KindGroup, "g9"); g == nil || g.Status != "targetOnly" || g.Copyable {
		t.Errorf("g9: want targetOnly + not copyable, got %+v", g)
	}
}

func TestMirrorCopyAddsOnlyMissingAndBringsTheTeamAlong(t *testing.T) {
	var calls []string
	sess := harness(t, mirrorFake(t, &calls))
	res, err := NewMirrorService(sess).Copy(MirrorRequest{
		Source: "src@contoso.com", Target: "dst@contoso.com",
		Kinds:   []string{KindGroup, KindRole, KindChannel, KindLicense},
		Confirm: "dst@contoso.com",
	})
	if err != nil {
		t.Fatal(err)
	}

	joined := strings.Join(calls, "\n")
	for _, want := range []string{
		"POST /users/dst-id/assignLicense",        // license the target lacked
		"POST /groups/g1/members/$ref",            // NOT wanted — asserted absent below
		"POST /directoryRoles/r1/members/$ref",    // missing admin role
		"POST /teams/t2/members",                  // team pulled in by its channel
		"POST /teams/t1/channels/c1/members",      // private channel in a team it already had
		"POST /teams/t2/channels/c2/members",      // private channel in the new team
	} {
		if want == "POST /groups/g1/members/$ref" {
			if strings.Contains(joined, want) {
				t.Errorf("a group the target already has must not be re-added:\n%s", joined)
			}
			continue
		}
		if !strings.Contains(joined, want) {
			t.Errorf("missing call %q in:\n%s", want, joined)
		}
	}
	// Teams were not selected, but t1 was already shared — it must not be re-added.
	if strings.Contains(joined, "POST /teams/t1/members") {
		t.Errorf("team the target already belongs to must not be re-added:\n%s", joined)
	}
	// Un-copyable groups are reported as failed steps, not silently dropped.
	var reported int
	for _, s := range res.Steps {
		if !s.OK && strings.Contains(s.Error, "not copyable") {
			reported++
		}
	}
	if reported != 2 {
		t.Errorf("want 2 un-copyable groups reported, got %d (%+v)", reported, res.Steps)
	}
	if res.OK {
		t.Error("a run that skipped un-copyable items must not report OK")
	}
}

func TestMirrorCopyNeedsTypedConfirmation(t *testing.T) {
	var calls []string
	sess := harness(t, mirrorFake(t, &calls))
	_, err := NewMirrorService(sess).Copy(MirrorRequest{
		Source: "src@contoso.com", Target: "dst@contoso.com",
		Kinds: []string{KindGroup}, Confirm: "wrong",
	})
	if err == nil {
		t.Fatal("copy without the typed target confirmation must fail")
	}
	if len(calls) != 0 {
		t.Errorf("nothing may be written before the confirmation passes: %v", calls)
	}
}
