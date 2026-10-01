package services

import (
	"net/http"
	"testing"
)

// The recommendations API has no v1.0 counterpart: both calls must go to the
// /beta version of whatever host the client is pointed at.
func TestRecommendationsUseBetaEndpoints(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		w.Write([]byte(`{"value":[{"id":"t_Microsoft.Identity.IAM.Insights.TurnOffPerUserMFA","priority":"medium"}]}`))
	})
	sec := NewSecurityService(sess)

	recs, err := sec.Recommendations()
	if err != nil {
		t.Fatal(err)
	}
	if len(recs) != 1 {
		t.Fatalf("want 1 recommendation, got %d", len(recs))
	}
	if _, err := sec.RecommendationImpacted("t_Microsoft.Identity.IAM.Insights.TurnOffPerUserMFA"); err != nil {
		t.Fatal(err)
	}

	want := []string{
		"GET /beta/directory/recommendations",
		"GET /beta/directory/recommendations/t_Microsoft.Identity.IAM.Insights.TurnOffPerUserMFA/impactedResources",
	}
	if len(calls) != len(want) {
		t.Fatalf("calls = %v", calls)
	}
	for i := range want {
		if calls[i] != want[i] {
			t.Errorf("call %d: want %q, got %q", i, want[i], calls[i])
		}
	}
}

// A 403 on recommendations must carry the permission hint so the UI can say
// exactly what to grant.
func TestRecommendationsForbiddenCarriesHint(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(403)
		w.Write([]byte(`{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}`))
	})
	_, err := NewSecurityService(sess).Recommendations()
	if err == nil {
		t.Fatal("want error")
	}
	oe, ok := err.(*OpError)
	if !ok {
		t.Fatalf("want *OpError, got %T: %v", err, err)
	}
	if oe.Status != 403 || oe.Hint != "DirectoryRecommendations.Read.All" {
		t.Errorf("envelope = %+v", oe)
	}
}
