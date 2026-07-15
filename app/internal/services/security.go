package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

// SecurityService — read-only tenant security review: Conditional Access
// policies and enterprise-app (service principal) consent overview.
// Requires Policy.Read.All (CA) and Application.Read.All (SPs + grants).
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
