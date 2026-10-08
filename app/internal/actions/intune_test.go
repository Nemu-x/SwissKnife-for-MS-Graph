package actions

import (
	"encoding/json"
	"io"
	"net/http"
	"strings"
	"testing"

	"swissknife-app/internal/engine"
)

// intuneGraph serves two configuration policies, one compliance policy, an
// assigned app and an empty settings catalog.
var denyApps bool

func intuneGraph(t *testing.T, posted *[]map[string]any) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		p := r.URL.Path
		switch {
		case r.Method == "POST" && strings.HasSuffix(p, "/assign"):
			var body map[string]any
			b, _ := io.ReadAll(r.Body)
			_ = json.Unmarshal(b, &body)
			*posted = append(*posted, body)
			w.WriteHeader(http.StatusNoContent)
		case p == "/users/ann@contoso.com/transitiveMemberOf/microsoft.graph.group":
			w.Write([]byte(`{"value":[{"id":"g1","displayName":"Sales"},{"id":"g2","displayName":"Contractors"}]}`))
		case p == "/groups/g1":
			w.Write([]byte(`{"id":"g1","displayName":"Sales"}`))
		case p == "/groups/g2":
			w.Write([]byte(`{"id":"g2","displayName":"Contractors"}`))
		case p == "/deviceManagement/deviceConfigurations":
			w.Write([]byte(`{"value":[
				{"id":"c1","displayName":"Wi-Fi","assignments":[
					{"target":{"@odata.type":"#microsoft.graph.groupAssignmentTarget","groupId":"g1"}},
					{"target":{"@odata.type":"#microsoft.graph.exclusionGroupAssignmentTarget","groupId":"g2"}},
					{"target":{"@odata.type":"#microsoft.graph.allDevicesAssignmentTarget","deviceAndAppManagementAssignmentFilterId":"f-corp","deviceAndAppManagementAssignmentFilterType":"include"}}]},
				{"id":"c2","displayName":"Old VPN","lastModifiedDateTime":"2024-01-01T00:00:00Z","assignments":[]}]}`))
		case p == "/deviceManagement/deviceConfigurations/c1/deviceStatuses":
			w.Write([]byte(`{"value":[{"deviceDisplayName":"LAPTOP-1","userPrincipalName":"ann@contoso.com","status":"conflict"}]}`))
		case p == "/deviceManagement/deviceCompliancePolicies":
			w.Write([]byte(`{"value":[{"id":"p1","displayName":"Baseline","assignments":[
				{"target":{"@odata.type":"#microsoft.graph.allLicensedUsersAssignmentTarget"}}]}]}`))
		case strings.HasSuffix(p, "/deviceStatuses"):
			w.Write([]byte(`{"value":[]}`))
		case p == "/deviceAppManagement/mobileApps" && denyApps:
			w.WriteHeader(http.StatusForbidden)
			w.Write([]byte(`{"error":{"code":"Forbidden","message":"no"}}`))
		case p == "/deviceAppManagement/mobileApps":
			w.Write([]byte(`{"value":[{"id":"a1","displayName":"Teams","assignments":[
				{"intent":"required","target":{"@odata.type":"#microsoft.graph.groupAssignmentTarget","groupId":"g1"}}]},
				{"id":"a2","displayName":"Unused app","assignments":[]}]}`))
		case strings.HasSuffix(p, "/configurationPolicies"):
			w.Write([]byte(`{"value":[]}`))
		default:
			t.Errorf("unexpected %s %s", r.Method, p)
		}
	}
}

func TestIntuneAssignedToUser(t *testing.T) {
	var posted []map[string]any
	e := securityHarness(t, intuneGraph(t, &posted), nil)
	res, err := e.Run(t.Context(), "intune.assignedTo", engine.Inputs{"user": "ann@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]string{}
	for _, r := range res.Rows {
		got[r["policy"]+"/"+r["via"]] = r["effect"]
	}
	want := map[string]string{
		"Wi-Fi/Sales": "included", "Wi-Fi/Contractors": "excluded", "Wi-Fi/allDevices": "included",
		"Baseline/allUsers": "included", "Teams/Sales": "required",
	}
	for k, v := range want {
		if got[k] != v {
			t.Errorf("%s = %q, want %q (rows %+v)", k, got[k], v, res.Rows)
		}
	}
	if _, err := e.Run(t.Context(), "intune.assignedTo", engine.Inputs{}); err == nil {
		t.Fatal("a user or a group is required")
	}
}

func TestIntuneUnassignedAndConflicts(t *testing.T) {
	var posted []map[string]any
	e := securityHarness(t, intuneGraph(t, &posted), nil)
	res, err := e.Run(t.Context(), "intune.unassigned", engine.Inputs{})
	if err != nil || len(res.Rows) != 1 || res.Rows[0]["policy"] != "Old VPN" {
		t.Fatalf("unassigned %+v %v", res, err)
	}
	res, err = e.Run(t.Context(), "intune.conflicts", engine.Inputs{})
	if err != nil || len(res.Rows) != 1 || res.Rows[0]["device"] != "LAPTOP-1" {
		t.Fatalf("conflicts %+v %v", res, err)
	}
}

func TestIntuneAssignKeepsTheOtherAssignments(t *testing.T) {
	var posted []map[string]any
	e := securityHarness(t, intuneGraph(t, &posted), nil)
	if _, err := e.Plan("intune.assign", engine.Inputs{"policy": "missing", "group": "g1"}); err == nil {
		t.Fatal("an unknown policy name must be refused")
	}
	// Removing Sales from Wi-Fi keeps the exclusion and the all-devices target.
	ch := planApply(t, e, "intune.assign", engine.Inputs{"policy": "wi-fi", "group": "g1", "op": "remove"})
	if ch.Op != "remove" || ch.Before != "Sales" {
		t.Fatalf("change %+v", ch)
	}
	list, _ := posted[0]["assignments"].([]any)
	if len(list) != 2 {
		t.Fatalf("assignments sent %+v", posted[0])
	}
	for _, a := range list {
		target := a.(map[string]any)["target"].(map[string]any)
		if strings.HasSuffix(target["@odata.type"].(string), "allDevicesAssignmentTarget") {
			if _, has := target["groupId"]; has {
				t.Fatalf("an all-devices target must not carry a group id: %+v", target)
			}
			// The assignment filter must survive the rebuilt list.
			if target["deviceAndAppManagementAssignmentFilterId"] != "f-corp" || target["deviceAndAppManagementAssignmentFilterType"] != "include" {
				t.Fatalf("assignment filter lost: %+v", target)
			}
		}
		if target["groupId"] == "g1" {
			t.Fatalf("Sales must be gone: %+v", list)
		}
	}
}

func TestIntuneAddingAnExcludedGroupIsRefused(t *testing.T) {
	var posted []map[string]any
	e := securityHarness(t, intuneGraph(t, &posted), nil)
	if _, err := e.Plan("intune.assign", engine.Inputs{"policy": "Wi-Fi", "group": "g2"}); err == nil {
		t.Fatal("including a group that is excluded must be refused")
	}
}

func TestIntuneReportsSayWhatTheyCouldNotRead(t *testing.T) {
	denyApps = true
	t.Cleanup(func() { denyApps = false })
	var posted []map[string]any
	e := securityHarness(t, intuneGraph(t, &posted), nil)
	res, err := e.Run(t.Context(), "intune.assignedTo", engine.Inputs{"group": "g1"})
	if err != nil {
		t.Fatal(err)
	}
	if res.Note == nil || res.Note.Key != "intuneDenied" || !strings.Contains(res.Note.Params["kinds"], "app") {
		t.Fatalf("note %+v", res.Note)
	}
}
