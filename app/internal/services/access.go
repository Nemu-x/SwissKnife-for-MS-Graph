package services

import (
	"net/url"
	"strings"

	"swissknife-app/internal/session"
)

// AccessService probes which features the connected app registration can use,
// so the UI can hide tabs the tenant has no permission for.
type AccessService struct {
	s *session.Session
}

func NewAccessService(s *session.Session) *AccessService { return &AccessService{s: s} }

// probe endpoints keyed by nav id. A feature is available only if the cheap
// probe call succeeds — any error (403 no permission, 400 "not onboarded", …)
// marks it unavailable so it can be hidden.
var accessProbes = map[string]string{
	"users":     "/users",
	"groups":    "/groups",
	"roles":     "/directoryRoles",
	"licensing": "/subscribedSkus",
	"devices":   "/devices",
	"intune":    "/deviceManagement/managedDevices",
	"apps":      "/applications",
	"security":  "/identity/conditionalAccess/policies",
	"audit":     "/auditLogs/signIns",
	"health":    "/admin/serviceAnnouncement/healthOverviews",
	"teams":     "/groups",
}

// Probe returns nav id → accessible. Runs cheap $top=1 GETs per feature, and
// additionally gates Intune on an actual Intune license being present (the
// endpoint can return an empty 200 even when the tenant has no Intune).
func (a *AccessService) Probe() map[string]bool {
	out := map[string]bool{}
	c, err := a.s.Client()
	if err != nil {
		return out
	}
	for id, path := range accessProbes {
		var sink map[string]any
		e := c.Get(a.s.Ctx(), path, url.Values{"$top": {"1"}}, &sink)
		out[id] = e == nil
	}

	// Intune: also require a licensed Intune service plan in the tenant.
	if out["intune"] {
		out["intune"] = a.hasServicePlan("INTUNE")
	}

	return out
}

// hasServicePlan reports whether any subscribed SKU includes a service plan
// whose name contains the given token (e.g. "INTUNE").
func (a *AccessService) hasServicePlan(token string) bool {
	c, err := a.s.Client()
	if err != nil {
		return false
	}
	var resp struct {
		Value []struct {
			ServicePlans []struct {
				ServicePlanName string `json:"servicePlanName"`
			} `json:"servicePlans"`
		} `json:"value"`
	}
	if err := c.Get(a.s.Ctx(), "/subscribedSkus", nil, &resp); err != nil {
		return false
	}
	token = strings.ToUpper(token)
	for _, sku := range resp.Value {
		for _, sp := range sku.ServicePlans {
			if strings.Contains(strings.ToUpper(sp.ServicePlanName), token) {
				return true
			}
		}
	}
	return false
}
