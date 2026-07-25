package services

import (
	"net/http"
	"strings"
	"testing"
)

// stepsByName filters a result's steps by name.
func stepsByName(steps []Step, name string) []Step {
	var out []Step
	for _, s := range steps {
		if s.Name == name {
			out = append(out, s)
		}
	}
	return out
}

// TestOffboardIntuneRetireAndWipe: each managed device gets its own step and
// the right Intune action endpoint.
func TestOffboardIntuneRetireAndWipe(t *testing.T) {
	for _, action := range []string{"retire", "wipe"} {
		var actions []string
		sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
			switch {
			case strings.HasSuffix(r.URL.Path, "/managedDevices"):
				w.Write([]byte(`{"value":[
					{"id":"d1","deviceName":"LAPTOP-01","operatingSystem":"Windows"},
					{"id":"d2","deviceName":"iPhone","operatingSystem":"iOS"}]}`))
			case strings.HasPrefix(r.URL.Path, "/deviceManagement/managedDevices/"):
				actions = append(actions, r.URL.Path)
				w.WriteHeader(204)
			default:
				w.Write([]byte(`{}`))
			}
		})
		pb := NewPlaybookService(sess)

		req := fullOffboardRequest()
		req.IntuneAction = action
		res, err := pb.Offboard(req)
		if err != nil {
			t.Fatal(err)
		}
		if len(actions) != 2 || !strings.HasSuffix(actions[0], "/d1/"+action) || !strings.HasSuffix(actions[1], "/d2/"+action) {
			t.Fatalf("%s: wrong action calls: %v", action, actions)
		}
		stepName := "Retire device"
		if action == "wipe" {
			stepName = "Wipe device"
		}
		devSteps := stepsByName(res.Steps, stepName)
		if len(devSteps) != 2 || !devSteps[0].OK || devSteps[0].Detail != "LAPTOP-01 (Windows)" {
			t.Fatalf("%s: device steps wrong: %+v", action, devSteps)
		}
	}
}

// TestOffboardRemovesMfaMethods: deletable methods each get a step; the
// password method is skipped.
func TestOffboardRemovesMfaMethods(t *testing.T) {
	var deletes []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.HasSuffix(r.URL.Path, "/authentication/methods"):
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.phoneAuthenticationMethod","id":"m1"},
				{"@odata.type":"#microsoft.graph.microsoftAuthenticatorAuthenticationMethod","id":"m2"},
				{"@odata.type":"#microsoft.graph.passwordAuthenticationMethod","id":"m3"}]}`))
		case r.Method == "DELETE" && strings.Contains(r.URL.Path, "/authentication/"):
			deletes = append(deletes, r.URL.Path)
			w.WriteHeader(204)
		default:
			w.Write([]byte(`{}`))
		}
	})
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.RemoveMfaMethods = true
	res, err := pb.Offboard(req)
	if err != nil {
		t.Fatal(err)
	}
	if len(deletes) != 2 {
		t.Fatalf("want 2 deletions (password skipped), got %v", deletes)
	}
	mfaSteps := stepsByName(res.Steps, "Remove MFA method")
	if len(mfaSteps) != 2 || mfaSteps[0].Detail != "phone" || mfaSteps[1].Detail != "microsoftAuthenticator" {
		t.Fatalf("mfa steps wrong: %+v", mfaSteps)
	}
}

// TestOffboardDeletesRegisteredDevices: device objects are deleted via /devices.
func TestOffboardDeletesRegisteredDevices(t *testing.T) {
	var deletes []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.HasSuffix(r.URL.Path, "/registeredDevices"):
			w.Write([]byte(`{"value":[
				{"@odata.type":"#microsoft.graph.device","id":"dev1","displayName":"DESKTOP-X"}]}`))
		case r.Method == "DELETE" && strings.HasPrefix(r.URL.Path, "/devices/"):
			deletes = append(deletes, r.URL.Path)
			w.WriteHeader(204)
		default:
			w.Write([]byte(`{}`))
		}
	})
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.DeleteRegisteredDevices = true
	res, err := pb.Offboard(req)
	if err != nil {
		t.Fatal(err)
	}
	if len(deletes) != 1 || deletes[0] != "/devices/dev1" {
		t.Fatalf("wrong deletes: %v", deletes)
	}
	devSteps := stepsByName(res.Steps, "Delete registered device")
	if len(devSteps) != 1 || !devSteps[0].OK || devSteps[0].Detail != "DESKTOP-X" {
		t.Fatalf("device steps wrong: %+v", devSteps)
	}
}

// TestOffboardSharedMailboxCheck: a user-purpose mailbox fails the pre-flight
// with a convert-first message; a shared one passes.
func TestOffboardSharedMailboxCheck(t *testing.T) {
	for purpose, wantOK := range map[string]bool{"user": false, "shared": true} {
		p := purpose
		sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
			if strings.HasSuffix(r.URL.Path, "/mailboxSettings") && r.Method == "GET" {
				w.Write([]byte(`{"userPurpose":"` + p + `"}`))
				return
			}
			w.Write([]byte(`{}`))
		})
		pb := NewPlaybookService(sess)

		res, err := pb.Offboard(fullOffboardRequest()) // RemoveAllLicenses is on
		if err != nil {
			t.Fatal(err)
		}
		checks := stepsByName(res.Steps, "Check mailbox type")
		if len(checks) != 1 {
			t.Fatalf("%s: want one check step, got %+v", purpose, res.Steps)
		}
		if checks[0].OK != wantOK {
			t.Fatalf("%s: check OK=%v want %v (%+v)", purpose, checks[0].OK, wantOK, checks[0])
		}
		if !wantOK && !strings.Contains(checks[0].Error, "convert to a shared mailbox") {
			t.Fatalf("user mailbox must explain conversion, got %q", checks[0].Error)
		}
		// The check must precede license removal in the report.
		licenseIdx, checkIdx := -1, -1
		for i, s := range res.Steps {
			if s.Name == "Remove licenses" {
				licenseIdx = i
			}
			if s.Name == "Check mailbox type" {
				checkIdx = i
			}
		}
		if checkIdx == -1 || licenseIdx == -1 || checkIdx > licenseIdx {
			t.Fatalf("%s: check must run before license removal (%d vs %d)", purpose, checkIdx, licenseIdx)
		}
	}
}
