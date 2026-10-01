package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

// SecurityService — read-only tenant security review: Conditional Access
// policies, enterprise-app (service principal) consent overview and Entra
// recommendations. Requires Policy.Read.All (CA), Application.Read.All (SPs +
// grants) and DirectoryRecommendations.Read.All (recommendations).
type SecurityService struct {
	s *session.Session
}

func NewSecurityService(s *session.Session) *SecurityService { return &SecurityService{s: s} }

// CAPolicies lists Conditional Access policies.
func (x *SecurityService) CAPolicies() ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(x.s.Ctx(), "/identity/conditionalAccess/policies", nil, 0)
}

// ServicePrincipals lists enterprise apps (service principals).
func (x *SecurityService) ServicePrincipals(search string, maxItems int) ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{
		"$top":    {"100"},
		"$select": {"id,appId,displayName,accountEnabled,servicePrincipalType,appOwnerOrganizationId,tags"},
	}
	if search != "" {
		params.Set("$filter", "startswith(displayName,'"+escapeODataLiteral(search)+"')")
	}
	return c.ListAll(x.s.Ctx(), "/servicePrincipals", params, maxItems)
}

// OAuthGrants lists the delegated permission grants (scopes) issued to a
// service principal — what the app can do on behalf of users.
func (x *SecurityService) OAuthGrants(spID string) ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(x.s.Ctx(), "/servicePrincipals/"+url.PathEscape(spID)+"/oauth2PermissionGrants", nil, 0)
}

// AppRoleAssignments lists the application permissions (app roles) granted to
// a service principal — what the app can do with no user present.
func (x *SecurityService) AppRoleAssignments(spID string) ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(x.s.Ctx(), "/servicePrincipals/"+url.PathEscape(spID)+"/appRoleAssignments", nil, 0)
}

// Recommendations lists Entra recommendations: what Microsoft's daily tenant
// analysis says should be fixed, each with insights, benefits, action steps,
// priority and status. The API exists only under /beta (no v1.0 counterpart
// as of 2026-09) and needs DirectoryRecommendations.Read.All. The API itself
// is license-free; individual recommendation types follow their feature's
// license — Identity Protection ones (sign-in / user risk policies) and most
// Identity Secure Score items appear only on Entra ID P2 tenants.
func (x *SecurityService) Recommendations() ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	out, err := c.ListAll(x.s.Ctx(), c.Beta("/directory/recommendations"), nil, 0)
	return out, wrapOpErr(err)
}

// RecommendationImpacted lists the directory objects a recommendation applies
// to (users, apps, service principals…). Tenant-level recommendations have
// none — the action plan then concerns the whole tenant.
func (x *SecurityService) RecommendationImpacted(id string) ([]json.RawMessage, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	path := c.Beta("/directory/recommendations/" + url.PathEscape(id) + "/impactedResources")
	// A recommendation can impact many directory objects; cap the crawl like
	// the snapshot collections do.
	out, err := c.ListAll(x.s.Ctx(), path, nil, snapshotListCap)
	return out, wrapOpErr(err)
}
