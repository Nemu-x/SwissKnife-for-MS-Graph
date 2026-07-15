package services

import (
	"encoding/json"
	"io"
	"net/http"
	"strings"
	"testing"
	"time"
)

// --- Offboarding "to the end" ---

// offboardHarness fakes every endpoint the full offboard sequence touches and
// records "METHOD path" for order assertions.
func offboardHarness(t *testing.T, calls *[]string, bodies map[string]string) http.HandlerFunc {
	t.Helper()
	return func(w http.ResponseWriter, r *http.Request) {
		key := r.Method + " " + r.URL.Path
		*calls = append(*calls, key)
		if b, _ := io.ReadAll(r.Body); len(b) > 0 {
			bodies[key] = string(b)
		}
		switch {
		case r.Method == "GET" && r.URL.Path == "/users/dep@contoso.com":
			w.Write([]byte(`{"id":"uid1"}`))
		case r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/memberOf"):
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.group","id":"g1","displayName":"Sales"},
				{"@odata.type":"#microsoft.graph.directoryRole","id":"r1","displayName":"Admins"}]}`))
		case r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/licenseDetails"):
			w.Write([]byte(`{"value":[{"skuId":"sku-1"}]}`))
		default:
			w.Write([]byte(`{}`))
		}
	}
}

func fullOffboardRequest() OffboardRequest {
	return OffboardRequest{
		Upn: "dep@contoso.com", Confirm: "dep@contoso.com",
		Block: true, RevokeSessions: true,
		Oof: true, ForwardTo: "boss@contoso.com", HideFromGal: true,
		CalendarTo: "boss@contoso.com", RemoveFromGroups: true,
		RemoveAllLicenses: true, Delete: false,
	}
}

func TestOffboardRunsStepsInOrderAndNeverDeletesUnlessAsked(t *testing.T) {
	var calls []string
	bodies := map[string]string{}
	sess := harness(t, offboardHarness(t, &calls, bodies))
	pb := NewPlaybookService(sess)

	res, err := pb.Offboard(fullOffboardRequest())
	if err != nil {
		t.Fatal(err)
	}
	if !res.OK {
		t.Fatalf("expected all steps OK, got %+v", res.Steps)
	}

	want := []string{
		"PATCH /users/dep@contoso.com",                               // block sign-in
		"POST /users/dep@contoso.com/revokeSignInSessions",           // revoke
		"PATCH /users/dep@contoso.com/mailboxSettings",               // OOF
		"POST /users/dep@contoso.com/mailFolders/inbox/messageRules", // forward
		"PATCH /users/dep@contoso.com",                               // hide from GAL
		"POST /users/dep@contoso.com/calendar/calendarPermissions",   // calendar
		"GET /users/dep@contoso.com",                                 // resolve id
		"GET /users/dep@contoso.com/memberOf",                        // list groups
		"DELETE /groups/g1/members/uid1/$ref",                        // remove from group
		"GET /users/dep@contoso.com/licenseDetails",                  // list licenses
		"POST /users/dep@contoso.com/assignLicense",                  // remove licenses
	}
	if len(calls) != len(want) {
		t.Fatalf("want %d calls, got %d: %v", len(want), len(calls), calls)
	}
	for i := range want {
		if calls[i] != want[i] {
			t.Errorf("call %d: want %q, got %q", i, want[i], calls[i])
		}
	}
	for _, c := range calls {
		if strings.HasPrefix(c, "DELETE /users/") {
			t.Fatalf("user was deleted although Delete=false: %v", calls)
		}
	}
	// Directory roles in memberOf must not be treated as groups.
	for _, c := range calls {
		if strings.Contains(c, "/groups/r1/") {
			t.Errorf("tried to remove membership from a directoryRole: %v", calls)
		}
	}
}

func TestOffboardDeleteRunsLastWhenRequested(t *testing.T) {
	var calls []string
	sess := harness(t, offboardHarness(t, &calls, map[string]string{}))
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.Delete = true
	if _, err := pb.Offboard(req); err != nil {
		t.Fatal(err)
	}
	last := calls[len(calls)-1]
	if last != "DELETE /users/dep@contoso.com" {
		t.Fatalf("delete must be the final step, got %q", last)
	}
}

func TestOffboardForwardAndOofBodies(t *testing.T) {
	var calls []string
	bodies := map[string]string{}
	sess := harness(t, offboardHarness(t, &calls, bodies))
	pb := NewPlaybookService(sess)

	if _, err := pb.Offboard(fullOffboardRequest()); err != nil {
		t.Fatal(err)
	}
	fwd := bodies["POST /users/dep@contoso.com/mailFolders/inbox/messageRules"]
	if !strings.Contains(fwd, "boss@contoso.com") || !strings.Contains(fwd, `"isEnabled":true`) {
		t.Errorf("forward rule body wrong: %s", fwd)
	}
	oof := bodies["PATCH /users/dep@contoso.com/mailboxSettings"]
	if !strings.Contains(oof, "alwaysEnabled") || !strings.Contains(oof, "internalReplyMessage") {
		t.Errorf("OOF body wrong: %s", oof)
	}
	gal := bodies["PATCH /users/dep@contoso.com"]
	if !strings.Contains(gal, `"showInAddressList":false`) {
		t.Errorf("hide-from-GAL body wrong: %s", gal)
	}
}

func TestOffboardRequiresTypedConfirm(t *testing.T) {
	called := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { called = true })
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.Confirm = "wrong"
	if _, err := pb.Offboard(req); err == nil {
		t.Fatal("expected confirm mismatch error")
	}
	if called {
		t.Error("request reached the server despite bad confirm")
	}
}

// --- Cleanup: version trimming ---

func TestTrimVersionsKeepsNewestAndCurrent(t *testing.T) {
	var deleted []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/versions"):
			// Graph returns versions newest-first.
			w.Write([]byte(`{"value":[{"id":"current"},{"id":"5.0"},{"id":"4.0"},{"id":"3.0"}]}`))
		case r.Method == "DELETE":
			deleted = append(deleted, r.URL.Path)
			w.WriteHeader(204)
		default:
			w.Write([]byte(`{}`))
		}
	})
	cl := NewCleanupService(sess)

	out, err := cl.TrimVersions("/drives/d/items/i", 2, "TRIM")
	if err != nil {
		t.Fatal(err)
	}
	// keep=2 keeps "current" and "5.0"; "4.0" and "3.0" go.
	if out["removed"] != 2 {
		t.Fatalf("want 2 removed, got %v", out["removed"])
	}
	want := []string{"/drives/d/items/i/versions/4.0", "/drives/d/items/i/versions/3.0"}
	if len(deleted) != 2 || deleted[0] != want[0] || deleted[1] != want[1] {
		t.Errorf("deleted wrong versions: %v", deleted)
	}
}

func TestTrimVersionsNeverDeletesCurrent(t *testing.T) {
	var deleted []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.Method {
		case "GET":
			w.Write([]byte(`{"value":[{"id":"2.0"},{"id":"current"},{"id":"1.0"}]}`))
		case "DELETE":
			deleted = append(deleted, r.URL.Path)
			w.WriteHeader(204)
		}
	})
	cl := NewCleanupService(sess)

	if _, err := cl.TrimVersions("/drives/d/items/i", 1, "TRIM"); err != nil {
		t.Fatal(err)
	}
	for _, d := range deleted {
		if strings.HasSuffix(d, "/current") {
			t.Fatalf("current version was deleted: %v", deleted)
		}
	}
}

func TestTrimVersionsRequiresConfirm(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	cl := NewCleanupService(sess)
	if _, err := cl.TrimVersions("/drives/d/items/i", 1, "trim"); err == nil {
		t.Fatal("lowercase confirm must be rejected")
	}
}

// --- TAP ---

func TestCreateTAPClampsLifetimeAndParsesBody(t *testing.T) {
	var body map[string]any
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		json.NewDecoder(r.Body).Decode(&body)
		w.Write([]byte(`{"temporaryAccessPass":"secret-pass","lifetimeInMinutes":60}`))
	})
	am := NewAuthMethodsService(sess)

	out, err := am.CreateTAP("new@contoso.com", 3, true) // 3 min is below Graph's minimum
	if err != nil {
		t.Fatal(err)
	}
	if body["lifetimeInMinutes"] != float64(60) {
		t.Errorf("lifetime not clamped up: %v", body["lifetimeInMinutes"])
	}
	if body["isUsableOnce"] != true {
		t.Errorf("isUsableOnce lost: %v", body)
	}
	if out["temporaryAccessPass"] != "secret-pass" {
		t.Errorf("pass not returned: %v", out)
	}

	if _, err := am.CreateTAP("new@contoso.com", 99999999, false); err != nil {
		t.Fatal(err)
	}
	if body["lifetimeInMinutes"] != float64(43200) {
		t.Errorf("lifetime not clamped down to 30 days: %v", body["lifetimeInMinutes"])
	}
}

// --- B2B invite ---

func TestInviteGuestBodyAndDefaults(t *testing.T) {
	var path string
	var body map[string]any
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		path = r.URL.Path
		json.NewDecoder(r.Body).Decode(&body)
		w.Write([]byte(`{"inviteRedeemUrl":"https://redeem"}`))
	})
	us := NewUsersService(sess)

	out, err := us.InviteGuest("guest@example.com", "", "", "", true)
	if err != nil {
		t.Fatal(err)
	}
	if path != "/invitations" {
		t.Errorf("wrong endpoint: %s", path)
	}
	if body["invitedUserEmailAddress"] != "guest@example.com" {
		t.Errorf("email missing: %v", body)
	}
	if body["inviteRedirectUrl"] != "https://myapps.microsoft.com" {
		t.Errorf("redirect default missing: %v", body)
	}
	if _, has := body["invitedUserDisplayName"]; has {
		t.Error("empty display name must be omitted")
	}
	if out["inviteRedeemUrl"] != "https://redeem" {
		t.Errorf("redeem url not returned: %v", out)
	}
}

// --- App secret rotation ---

func TestAddSecretClampsMonthsAndBuildsBody(t *testing.T) {
	var body struct {
		PasswordCredential struct {
			DisplayName string    `json:"displayName"`
			EndDateTime time.Time `json:"endDateTime"`
		} `json:"passwordCredential"`
	}
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		json.NewDecoder(r.Body).Decode(&body)
		w.Write([]byte(`{"secretText":"s3cr3t"}`))
	})
	apps := NewAppsService(sess)

	out, err := apps.AddSecret("obj-1", "", 99) // out of range -> default 6 months
	if err != nil {
		t.Fatal(err)
	}
	if out["secretText"] != "s3cr3t" {
		t.Errorf("secret not returned: %v", out)
	}
	if body.PasswordCredential.DisplayName == "" {
		t.Error("default display name missing")
	}
	lo := time.Now().AddDate(0, 6, -2)
	hi := time.Now().AddDate(0, 6, 2)
	if body.PasswordCredential.EndDateTime.Before(lo) || body.PasswordCredential.EndDateTime.After(hi) {
		t.Errorf("endDateTime not ~6 months out: %v", body.PasswordCredential.EndDateTime)
	}
}

func TestAddSecretBlockedInReadOnly(t *testing.T) {
	called := false
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) { called = true })
	sess.SetReadOnly(true)
	apps := NewAppsService(sess)
	if _, err := apps.AddSecret("obj-1", "x", 6); err == nil {
		t.Fatal("expected read-only error")
	}
	if called {
		t.Error("write reached server in read-only mode")
	}
}

// --- Security review (read-only endpoints) ---

func TestSecurityServiceEndpoints(t *testing.T) {
	var paths []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		paths = append(paths, r.URL.Path)
		w.Write([]byte(`{"value":[{"id":"1"}]}`))
	})
	sec := NewSecurityService(sess)

	if _, err := sec.CAPolicies(); err != nil {
		t.Fatal(err)
	}
	if _, err := sec.OAuthGrants("sp1"); err != nil {
		t.Fatal(err)
	}
	if _, err := sec.AppRoleAssignments("sp1"); err != nil {
		t.Fatal(err)
	}
	want := []string{
		"/identity/conditionalAccess/policies",
		"/servicePrincipals/sp1/oauth2PermissionGrants",
		"/servicePrincipals/sp1/appRoleAssignments",
	}
	for i, p := range want {
		if paths[i] != p {
			t.Errorf("endpoint %d: want %s, got %s", i, p, paths[i])
		}
	}
}
