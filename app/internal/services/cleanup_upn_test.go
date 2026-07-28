package services

import (
	"net/http"
	"strings"
	"testing"
)

// A UPN with an apostrophe used to be pasted straight into /users/{upn}/drive,
// where Graph's OData parser misread the following segments and answered
// "400 Request_BadRequest: Unexpected segment DynamicPathSegment". The scan must
// resolve the UPN to an object id first.
func TestVersionScanAddressesTheUserByObjectID(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch {
		case r.URL.Path == "/users/o'brien@contoso.com":
			w.Write([]byte(`{"id":"11111111-2222-3333-4444-555555555555"}`))
		case strings.HasSuffix(r.URL.Path, "/drive/root/children"):
			w.Write([]byte(`{"value":[]}`))
		default:
			w.Write([]byte(`{}`))
		}
	})

	if _, err := NewCleanupService(sess).FindVersionBloat("user", "o'brien@contoso.com", 2, 100); err != nil {
		t.Fatal(err)
	}

	joined := strings.Join(calls, "\n")
	if !strings.Contains(joined, "GET /users/11111111-2222-3333-4444-555555555555/drive/root/children") {
		t.Errorf("the walk must run against the object id, got:\n%s", joined)
	}
	for _, c := range calls {
		if strings.Contains(c, "o'brien@contoso.com/drive") {
			t.Errorf("a UPN must never appear in a deep drive path: %s", c)
		}
	}
}

// An object id passed in as the owner must not cost an extra lookup.
func TestVersionScanPassesAGuidThrough(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		w.Write([]byte(`{"value":[]}`))
	})

	id := "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
	if _, err := NewCleanupService(sess).FindVersionBloat("user", id, 2, 100); err != nil {
		t.Fatal(err)
	}
	for _, c := range calls {
		if c == "GET /users/"+id {
			t.Errorf("a GUID owner must not be resolved again: %v", calls)
		}
	}
}
