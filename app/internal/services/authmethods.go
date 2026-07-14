package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

// AuthMethodsService — Entra authentication methods (MFA) management.
type AuthMethodsService struct {
	s *session.Session
}

func NewAuthMethodsService(s *session.Session) *AuthMethodsService { return &AuthMethodsService{s: s} }

// methodTypeToSegment maps a method's @odata.type to its type-specific
// collection segment used for deletion. Password methods are not deletable.
var methodTypeToSegment = map[string]string{
	"#microsoft.graph.phoneAuthenticationMethod":                   "phoneMethods",
	"#microsoft.graph.microsoftAuthenticatorAuthenticationMethod":  "microsoftAuthenticatorMethods",
	"#microsoft.graph.softwareOathAuthenticationMethod":            "softwareOathMethods",
	"#microsoft.graph.fido2AuthenticationMethod":                   "fido2Methods",
	"#microsoft.graph.windowsHelloForBusinessAuthenticationMethod": "windowsHelloForBusinessMethods",
	"#microsoft.graph.emailAuthenticationMethod":                   "emailMethods",
	"#microsoft.graph.temporaryAccessPassAuthenticationMethod":     "temporaryAccessPassMethods",
}

// List returns the user's registered authentication methods.
func (a *AuthMethodsService) List(user string) ([]json.RawMessage, error) {
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(a.s.Ctx(), "/users/"+url.PathEscape(user)+"/authentication/methods", nil, 0)
}

// ResetMFA removes all non-password authentication methods so the user must
// re-register MFA. Destructive: requires typed confirm on the UPN.
func (a *AuthMethodsService) ResetMFA(user, confirm string) (map[string]any, error) {
	if err := a.s.GuardDestructive(user, confirm); err != nil {
		return nil, err
	}
	c, err := a.s.Client()
	if err != nil {
		return nil, err
	}
	methods, err := a.List(user)
	if err != nil {
		return nil, err
	}
	removed := 0
	failures := map[string]string{}
	for _, raw := range methods {
		var m struct {
			ID   string `json:"id"`
			Type string `json:"@odata.type"`
		}
		if json.Unmarshal(raw, &m) != nil {
			continue
		}
		seg, ok := methodTypeToSegment[m.Type]
		if !ok {
			continue // password or non-deletable
		}
		path := "/users/" + url.PathEscape(user) + "/authentication/" + seg + "/" + url.PathEscape(m.ID)
		if derr := c.Delete(a.s.Ctx(), path); derr != nil {
			failures[m.Type] = derr.Error()
			continue
		}
		removed++
	}
	a.s.Record("authMethods.resetMFA", user, "removed="+strconv.Itoa(removed), err)
	return map[string]any{"removed": removed, "failures": failures}, nil
}
