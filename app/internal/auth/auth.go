// Package auth acquires Graph tokens via azidentity.
// The token lives only here and in graphapi.Client; it is never exposed outward (ADR-002).
package auth

import (
	"context"
	"sync"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore"
	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
)

// Resource identifiers accepted by TokenFor (ADR-008). Tokens are requested
// with the resource's /.default scope.
const (
	ResourceGraph    = "https://graph.microsoft.com"
	ResourceExchange = "https://outlook.office365.com"
	ResourceTeams    = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239"
)

// Mode is the profile authentication mode.
type Mode string

const (
	ModeClientSecret Mode = "client_secret" // app-only
	ModeDeviceCode   Mode = "device_code"   // delegated
	// ModeClientCertificate is app-only with a certificate (PFX file).
	ModeClientCertificate Mode = "client_certificate"
)

// DeviceCodePrompt is invoked when the user must enter a code at microsoft.com/devicelogin.
type DeviceCodePrompt func(verificationURL, userCode, message string)

// TokenProvider implements graphapi.TokenSource on top of azcore.TokenCredential
// with per-resource caching and auto-refresh (proactively, 2 minutes before
// expiry).
type TokenProvider struct {
	cred azcore.TokenCredential

	mu     sync.Mutex
	tokens map[string]azcore.AccessToken
}

func NewClientSecret(tenantID, clientID, clientSecret string) (*TokenProvider, error) {
	cred, err := azidentity.NewClientSecretCredential(tenantID, clientID, clientSecret, nil)
	if err != nil {
		return nil, err
	}
	return &TokenProvider{cred: cred, tokens: map[string]azcore.AccessToken{}}, nil
}

func NewDeviceCode(tenantID, clientID string, prompt DeviceCodePrompt) (*TokenProvider, error) {
	cred, err := azidentity.NewDeviceCodeCredential(&azidentity.DeviceCodeCredentialOptions{
		TenantID: tenantID,
		ClientID: clientID,
		UserPrompt: func(ctx context.Context, dc azidentity.DeviceCodeMessage) error {
			if prompt != nil {
				prompt(dc.VerificationURL, dc.UserCode, dc.Message)
			}
			return nil
		},
	})
	if err != nil {
		return nil, err
	}
	return &TokenProvider{cred: cred, tokens: map[string]azcore.AccessToken{}}, nil
}

// Token implements graphapi.TokenSource (Microsoft Graph).
func (p *TokenProvider) Token(ctx context.Context) (string, error) {
	return p.TokenFor(ctx, ResourceGraph)
}

// TokenFor returns a token for the resource (one of the Resource* constants),
// cached per resource.
func (p *TokenProvider) TokenFor(ctx context.Context, resource string) (string, error) {
	p.mu.Lock()
	defer p.mu.Unlock()

	if t, ok := p.tokens[resource]; ok && time.Until(t.ExpiresOn) > 2*time.Minute {
		return t.Token, nil
	}

	tok, err := p.cred.GetToken(ctx, policy.TokenRequestOptions{Scopes: []string{resource + "/.default"}})
	if err != nil {
		return "", err
	}
	p.tokens[resource] = tok
	return tok.Token, nil
}
