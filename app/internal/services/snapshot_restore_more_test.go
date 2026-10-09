package services

import (
	"encoding/json"
	"io"
	"net/http"
	"strings"
	"testing"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/pwsh"
)

// Intune policies restore their settings and assignments; a deleted one is
// recreated, then assigned again. Everything goes in one "all" preview, one
// step per section.
func TestRestoreIntunePoliciesWithAssignments(t *testing.T) {
	type call struct{ method, path, body string }
	var calls []call
	notFound := func(w http.ResponseWriter) {
		w.WriteHeader(http.StatusNotFound)
		_, _ = w.Write([]byte(`{"error":{"code":"ResourceNotFound","message":"gone"}}`))
	}
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		calls = append(calls, call{r.Method, r.URL.Path, string(b)})
		switch {
		case r.Method == "GET" && r.URL.Path == "/deviceManagement/deviceConfigurations":
			_, _ = w.Write([]byte(`{"value":[{"id":"c1","displayName":"Win baseline"}]}`))
		case r.Method == "GET" && r.URL.Path == "/deviceManagement/deviceConfigurations/c1":
			_, _ = w.Write([]byte(`{"@odata.type":"#microsoft.graph.windows10GeneralConfiguration","id":"c1","displayName":"Win baseline",
				"passwordRequired":false,"wifiKey":"set","version":5,"assignments":[]}`))
		case r.Method == "GET" && strings.HasPrefix(r.URL.Path, "/groups/") && r.URL.Path != "/groups/gone":
			_, _ = w.Write([]byte(`{"id":"x"}`))
		case r.Method == "GET" && r.URL.Path == "/deviceManagement/deviceCompliancePolicies":
			_, _ = w.Write([]byte(`{"value":[]}`))
		case r.Method == "GET":
			notFound(w)
		case r.Method == "POST" && r.URL.Path == "/deviceManagement/deviceConfigurations":
			_, _ = w.Write([]byte(`{"id":"new2"}`))
		case r.Method == "POST" && r.URL.Path == "/deviceManagement/deviceCompliancePolicies":
			_, _ = w.Write([]byte(`{"id":"new3"}`))
		default:
			w.WriteHeader(http.StatusNoContent)
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sess.SetIdentity("tenant-a", true)
	group := func(id string) []any {
		return []any{map[string]any{"id": "a-" + id, "target": map[string]any{"@odata.type": "#microsoft.graph.groupAssignmentTarget", "groupId": id}}}
	}
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"intuneConfigurations": {
			{"@odata.type": "#microsoft.graph.windows10GeneralConfiguration", "id": "c1", "displayName": "Win baseline",
				"passwordRequired": true, "wifiKey": nil, "version": 3, "assignments": append(group("g1"), group("gone")...)},
			{"@odata.type": "#microsoft.graph.iosGeneralDeviceConfiguration", "id": "c2", "displayName": "iOS", "assignments": group("g2")},
		},
		"intuneCompliancePolicies": {
			{"@odata.type": "#microsoft.graph.windows10CompliancePolicy", "id": "p1", "displayName": "Win compliance", "assignments": group("g3"),
				"scheduledActionsForRule": []any{map[string]any{"id": "r1", "ruleName": "PasswordRequired", "scheduledActionConfigurations": []any{
					map[string]any{"id": "s1", "actionType": "block", "gracePeriodHours": 72}}}}},
		},
	})
	e := NewEngine(sess)
	p, err := e.Plan("config.restore", map[string]string{"snapshot": id, "section": "all"})
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]string{}
	for _, c := range p.Changes {
		got[c.Target] = c.Op + ":" + c.Before + ":" + c.Step
		if c.Target == "Win baseline" && c.Note != "assignmentGroupGone" {
			t.Errorf("a deleted assignment group is reported: %+v", c)
		}
	}
	want := map[string]string{
		"Win baseline":   "set:assignments, passwordRequired:snapshot.section.intuneConfigurations",
		"iOS":            "add::snapshot.section.intuneConfigurations",
		"Win compliance": "add::snapshot.section.intuneCompliancePolicies",
	}
	for k, v := range want {
		if got[k] != v {
			t.Errorf("%s = %q, want %q", k, got[k], v)
		}
	}
	if p.ConfirmTarget != "restore 3" {
		t.Fatalf("confirm %q", p.ConfirmTarget)
	}
	if _, err := e.Apply(p.ID, "restore 3"); err != nil {
		t.Fatal(err)
	}
	bodies := map[string]string{}
	for _, c := range calls {
		if c.method != "GET" {
			bodies[c.method+" "+c.path] = c.body
		}
	}
	patch := bodies["PATCH /deviceManagement/deviceConfigurations/c1"]
	if !strings.Contains(patch, `"passwordRequired":true`) || !strings.Contains(patch, "@odata.type") ||
		strings.Contains(patch, "wifiKey") || strings.Contains(patch, "version") || strings.Contains(patch, `"id"`) {
		t.Fatalf("patch body %s", patch)
	}
	for path, g := range map[string]string{
		"POST /deviceManagement/deviceConfigurations/c1/assign":       "g1",
		"POST /deviceManagement/deviceConfigurations/new2/assign":     "g2",
		"POST /deviceManagement/deviceCompliancePolicies/new3/assign": "g3",
	} {
		if !strings.Contains(bodies[path], `"groupId":"`+g+`"`) {
			t.Errorf("%s: %q", path, bodies[path])
		}
	}
	if b := bodies["POST /deviceManagement/deviceCompliancePolicies"]; !strings.Contains(b, `"gracePeriodHours":72`) || strings.Contains(b, `"s1"`) {
		t.Fatalf("a compliance policy is recreated with the snapshot's non-compliance actions: %s", b)
	}
	if strings.Contains(bodies["POST /deviceManagement/deviceConfigurations/c1/assign"], "gone") {
		t.Fatal("a deleted group is not assigned")
	}
}

type restorePS struct{ calls []string }

func (r *restorePS) Invoke(_ engine.Env, _, cmdlet string, params map[string]any, _ ...string) ([]json.RawMessage, error) {
	b, _ := json.Marshal(params)
	r.calls = append(r.calls, cmdlet+" "+string(b))
	switch cmdlet {
	case "Get-TransportRule":
		return []json.RawMessage{json.RawMessage(`{"Guid":"guid1","Name":"Block exe","State":"Disabled","Mode":"Enforce","Priority":0}`)}, nil
	case "Get-CsTeamsMeetingPolicy":
		return []json.RawMessage{json.RawMessage(`{"Identity":"Global","AllowCloudRecording":true}`)}, nil
	case "Get-OrganizationConfig":
		return []json.RawMessage{json.RawMessage(`{"Name":"contoso","AuditDisabled":true,"FocusedInboxOn":true}`)}, nil
	}
	return nil, nil
}

// Exchange and Teams settings go back through their Set- cmdlets; a deleted
// Teams policy is recreated, a deleted transport rule (whose conditions the
// snapshot does not keep) is reported, not invented.
func TestRestorePowerShellSections(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { _, _ = w.Write([]byte(`{}`)) })
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sess.SetIdentity("tenant-a", true)
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"exchangeOrganizationConfig": {{"id": "contoso", "Name": "contoso", "AuditDisabled": false, "FocusedInboxOn": true, "DefaultAuthenticationPolicy": nil}},
		"transportRules": {
			{"id": "guid1", "Guid": "guid1", "Name": "Block exe", "State": "Enabled", "Mode": "Enforce", "Priority": 0},
			{"id": "guid2", "Guid": "guid2", "Name": "Gone rule", "State": "Enabled"},
		},
		"teamsMeetingPolicies": {{"id": "Tag:NoRec", "Identity": "Tag:NoRec", "AllowCloudRecording": false}},
	})
	ps := &restorePS{}
	e := engine.New(sess, engine.GraphProvider{}, readyProvider{pwsh.BackendExchangePS}, readyProvider{pwsh.BackendTeamsPS})
	e.PS = ps
	e.Register(snapshotRestoreAction(sess))
	engines.Store(sess, e)
	t.Cleanup(func() { engines.Delete(sess) })

	p, err := e.Plan("config.restore", map[string]string{"snapshot": id, "section": "all"})
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]string{}
	for _, c := range p.Changes {
		got[c.Target] = c.Op + ":" + c.Before + ":" + c.Note
	}
	want := map[string]string{
		"contoso":   "set:AuditDisabled:",
		"Block exe": "set:State:",
		"Gone rule": "none::cannotRecreate",
		"Tag:NoRec": "add::",
	}
	for k, v := range want {
		if got[k] != v {
			t.Errorf("%s = %q, want %q", k, got[k], v)
		}
	}
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil {
		t.Fatal(err)
	}
	joined := strings.Join(ps.calls, "\n")
	for _, c := range []string{
		`Set-OrganizationConfig {"AuditDisabled":false}`,
		`Enable-TransportRule {"Confirm":false,"Identity":"guid1"}`,
		`New-CsTeamsMeetingPolicy {"AllowCloudRecording":false,"Identity":"NoRec"}`,
	} {
		if !strings.Contains(joined, c) {
			t.Errorf("missing call %s in\n%s", c, joined)
		}
	}
	if strings.Contains(joined, "Set-TransportRule") {
		t.Errorf("only the state changed: %s", joined)
	}
}

// Restoring everything recreates a deleted named location first, then
// points the policy that used it at its new id.
func TestRestoreAllRemapsRecreatedLocations(t *testing.T) {
	created := false
	var patch string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		switch {
		case r.Method == "GET" && r.URL.Path == "/identity/conditionalAccess/namedLocations":
			if created {
				_, _ = w.Write([]byte(`{"value":[{"id":"l1new","displayName":"Office"}]}`))
				return
			}
			_, _ = w.Write([]byte(`{"value":[]}`))
		case r.Method == "GET" && r.URL.Path == "/identity/conditionalAccess/policies":
			_, _ = w.Write([]byte(`{"value":[{"id":"pol1","displayName":"Office only"}]}`))
		case r.Method == "GET" && r.URL.Path == "/identity/conditionalAccess/policies/pol1":
			_, _ = w.Write([]byte(`{"id":"pol1","displayName":"Office only","state":"enabled","conditions":{}}`))
		case r.Method == "GET":
			w.WriteHeader(http.StatusNotFound)
			_, _ = w.Write([]byte(`{"error":{"code":"ResourceNotFound","message":"gone"}}`))
		case r.Method == "POST":
			created = true
			_, _ = w.Write([]byte(`{"id":"l1new"}`))
		case r.Method == "PATCH":
			patch = string(b)
			w.WriteHeader(http.StatusNoContent)
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sess.SetIdentity("tenant-a", true)
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"namedLocations": {{"@odata.type": "#microsoft.graph.ipNamedLocation", "id": "l1", "displayName": "Office"}},
		"conditionalAccessPolicies": {{"id": "pol1", "displayName": "Office only", "state": "enabled",
			"conditions": map[string]any{"locations": map[string]any{"includeLocations": []any{"l1"}}}}},
	})
	e := NewEngine(sess)
	p, err := e.Plan("config.restore", map[string]string{"snapshot": id, "section": "all"})
	if err != nil {
		t.Fatal(err)
	}
	if len(p.Changes) != 2 || p.Changes[0].Target != "Office" || p.Changes[0].Op != "add" || p.Changes[1].StepNo != 2 {
		t.Fatalf("plan %+v", p.Changes)
	}
	if _, err := e.Apply(p.ID, p.ConfirmTarget); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(patch, `"l1new"`) || strings.Contains(patch, `"l1"`) {
		t.Fatalf("the policy must point at the recreated location: %s", patch)
	}
}

// The typed "restore N" lands on a real change even when the first row is a
// section that could not be read.
func TestRestoreConfirmationSurvivesASkippedFirstSection(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/identity/conditionalAccess/namedLocations":
			_, _ = w.Write([]byte(`{"value":[]}`))
		case strings.HasPrefix(r.URL.Path, "/identity/conditionalAccess/namedLocations/"):
			w.WriteHeader(http.StatusForbidden)
			_, _ = w.Write([]byte(`{"error":{"code":"Forbidden","message":"no"}}`))
		case r.URL.Path == "/identity/conditionalAccess/policies":
			_, _ = w.Write([]byte(`{"value":[]}`))
		default:
			w.WriteHeader(http.StatusNotFound)
			_, _ = w.Write([]byte(`{"error":{"code":"ResourceNotFound","message":"gone"}}`))
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	sess.SetIdentity("tenant-a", true)
	id := savedSnapshot(t, dir, map[string][]map[string]any{
		"namedLocations":            {{"id": "l1", "displayName": "Office"}},
		"conditionalAccessPolicies": {{"id": "p1", "displayName": "MFA", "state": "enabled"}},
	})
	e := NewEngine(sess)
	p, err := e.Plan("config.restore", map[string]string{"snapshot": id, "section": "all"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Note != "sectionSkipped" || p.ConfirmTarget != "restore 1" {
		t.Fatalf("plan %+v confirm %q", p.Changes, p.ConfirmTarget)
	}
}
