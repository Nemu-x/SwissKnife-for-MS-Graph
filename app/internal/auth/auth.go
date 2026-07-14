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

var graphScopes = []string{"https://graph.microsoft.com/.default"}

// Mode is the profile authentication mode.
type Mode string

const (
	ModeClientSecret Mode = "client_secret" // app-only
	ModeDeviceCode   Mode = "device_code"   // delegated
)

// DeviceCodePrompt is invoked when the user must enter a code at microsoft.com/devicelogin.
type DeviceCodePrompt func(verificationURL, userCode, message string)

// TokenProvider implements graphapi.TokenSource on top of azcore.TokenCredential
// with caching and auto-refresh (proactively, 2 minutes before expiry).
type TokenProvider struct {
	cred azcore.TokenCredential

	mu    sync.Mutex
	token azcore.AccessToken
}

func NewClientSecret(tenantID, clientID, clientSecret string) (*TokenProvider, error) {
	cred, err := azidentity.NewClientSecretCredential(tenantID, clientID, clientSecret, nil)
	if err != nil {
		return nil, err
	}
	return &TokenProvider{cred: cred}, nil
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
	return &TokenProvider{cred: cred}, nil
}

// Token implements graphapi.TokenSource.
func (p *TokenProvider) Token(ctx context.Context) (string, error) {
	p.mu.Lock()
	defer p.mu.Unlock()

	if p.token.Token != "" && time.Until(p.token.ExpiresOn) > 2*time.Minute {
		return p.token.Token, nil
	}

	tok, err := p.cred.GetToken(ctx, policy.TokenRequestOptions{Scopes: graphScopes})
	if err != nil {
		return "", err
	}
	p.token = tok
	return tok.Token, nil
}
