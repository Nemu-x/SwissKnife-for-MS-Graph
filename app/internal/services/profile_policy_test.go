package services

import (
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

func TestProfilePolicySurvivesProfileEditsAndValidates(t *testing.T) {
	dir := t.TempDir()
	c := NewConnectService(session.New(auditlog.New(dir)), secrets.NewStoreAt(dir))
	p, err := c.SaveProfile(secrets.Profile{Name: "helpdesk", TenantID: "t", ClientID: "c", AuthMode: "device_code"}, "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := c.SetProfilePolicy(p.ID, session.Policy{MaxDanger: "admin"}); err == nil {
		t.Fatal("unknown ceiling must be refused")
	}
	got, err := c.SetProfilePolicy(p.ID, session.Policy{MaxDanger: "write", AllowedGroups: []string{" g1 ", "g1", ""},
		GroupLabels: map[string]string{"g1": "Helpdesk", "gone": "x"}})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Policy.AllowedGroups) != 1 || got.Policy.GroupLabels["gone"] != "" {
		t.Fatalf("policy %+v", got.Policy)
	}
	// The profile form does not send the policy: an edit must keep it.
	p.Name = "helpdesk 2"
	p.Policy = nil
	if _, err := c.SaveProfile(p, ""); err != nil {
		t.Fatal(err)
	}
	list, _ := c.Profiles()
	if len(list) != 1 || list[0].Policy == nil || list[0].Policy.MaxDanger != "write" {
		t.Fatalf("policy lost on edit: %+v", list)
	}
	// Clearing everything removes the policy.
	got, _ = c.SetProfilePolicy(p.ID, session.Policy{})
	if got.Policy != nil {
		t.Fatalf("empty policy must be dropped: %+v", got.Policy)
	}
}
