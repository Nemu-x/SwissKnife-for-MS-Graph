package actions

import (
	"net/http"
	"strings"
	"testing"

	"swissknife-app/internal/engine"
)

func TestInactiveAndLicenseWaste(t *testing.T) {
	var filter string
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/users":
			filter = r.URL.Query().Get("$filter")
			w.Write([]byte(`{"value":[
				{"userPrincipalName":"old@contoso.com","accountEnabled":true,"assignedLicenses":[{"skuId":"s1"}],"signInActivity":{"lastSignInDateTime":"2026-01-02T00:00:00Z"}},
				{"userPrincipalName":"free@contoso.com","accountEnabled":false,"assignedLicenses":[],"signInActivity":{"lastSignInDateTime":"2025-12-01T00:00:00Z"}}]}`))
		case "/subscribedSkus":
			w.Write([]byte(`{"value":[{"skuId":"s1","skuPartNumber":"ENTERPRISEPACK"}]}`))
		}
	}, nil)
	res, err := e.Run(t.Context(), "report.inactiveUsers", engine.Inputs{"days": "30"})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.HasPrefix(filter, "signInActivity/lastSignInDateTime le ") || len(res.Rows) != 2 || res.Rows[0]["user"] != "free@contoso.com" {
		t.Fatalf("filter %q rows %+v", filter, res.Rows)
	}
	res, err = e.Run(t.Context(), "report.licenseWaste", engine.Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != 1 || res.Rows[0]["licenses"] != "ENTERPRISEPACK" {
		t.Fatalf("license waste rows %+v", res.Rows)
	}
}

func TestMfaStatusPutsAdminsFirstAndSkipsGuests(t *testing.T) {
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"value":[
			{"userPrincipalName":"a@contoso.com","isAdmin":false,"methodsRegistered":[],"userType":"member"},
			{"userPrincipalName":"g@ext.example","isAdmin":false,"userType":"guest"},
			{"userPrincipalName":"root@contoso.com","isAdmin":true,"methodsRegistered":["email"],"userType":"member"}]}`))
	}, nil)
	res, err := e.Run(t.Context(), "report.mfaStatus", engine.Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != 2 || res.Rows[0]["user"] != "root@contoso.com" || res.Rows[0]["admin"] != "yes" {
		t.Fatalf("rows %+v", res.Rows)
	}
}

func TestPrivilegedRolesKinds(t *testing.T) {
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == "/directoryRoles" {
			w.Write([]byte(`{"value":[{"id":"r1","displayName":"Global Administrator"}]}`))
			return
		}
		w.Write([]byte(`{"value":[
			{"@odata.type":"#microsoft.graph.user","userPrincipalName":"root@contoso.com"},
			{"@odata.type":"#microsoft.graph.servicePrincipal","displayName":"Backup app"}]}`))
	}, nil)
	res, err := e.Run(t.Context(), "report.privilegedRoles", engine.Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != 2 || res.Rows[1]["kind"] != "app" || res.Rows[1]["member"] != "Backup app" {
		t.Fatalf("rows %+v", res.Rows)
	}
}

func TestMailboxSizesParsesTheUsageCSV(t *testing.T) {
	e := securityHarness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte("\xEF\xBB\xBFReport Refresh Date,User Principal Name,Item Count,Storage Used (Byte),Prohibit Send/Receive Quota (Byte),Last Activity Date\n" +
			"2026-10-07,small@contoso.com,10,1024,107374182400,2026-10-06\n" +
			"2026-10-07,big@contoso.com,90000,96636764160,107374182400,2026-10-07\n"))
	}, nil)
	res, err := e.Run(t.Context(), "report.mailboxSizes", engine.Inputs{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Rows) != 2 || res.Rows[0]["mailbox"] != "big@contoso.com" || res.Rows[0]["full"] != "90%" || res.Rows[0]["used"] != "90.0 GB" {
		t.Fatalf("rows %+v", res.Rows)
	}
}

func TestExchangeReports(t *testing.T) {
	fake := &fakePS{answers: map[string]string{
		"Get-Mailbox":           `[{"UserPrincipalName":"ann@contoso.com","ForwardingSmtpAddress":"smtp:x@evil.example","DeliverToMailboxAndForward":false}]`,
		"Get-MailboxStatistics": `[{"TotalItemSize":"1.2 GB (1,288,490,188 bytes)","ItemCount":4200,"LastUserActionTime":null}]`,
	}}
	e := securityHarness(t, graphUsers, fake)
	res, err := e.Run(t.Context(), "report.mailboxForwarding", engine.Inputs{})
	if err != nil || len(res.Rows) != 1 || res.Rows[0]["forwardsTo"] != "x@evil.example" || res.Rows[0]["keepsCopy"] != "no" {
		t.Fatalf("forwarding %+v %v", res, err)
	}
	if !strings.Contains(fake.calls[0].Params["Filter"].(string), "ForwardingSmtpAddress -ne $null") {
		t.Fatalf("filter %+v", fake.calls[0].Params)
	}
	res, err = e.Run(t.Context(), "mailbox.statistics", engine.Inputs{"mailbox": "ann@contoso.com"})
	if err != nil || res.Rows[1]["value"] != "4200" || res.Rows[4]["value"] != "" {
		t.Fatalf("statistics %+v %v", res, err)
	}
}
