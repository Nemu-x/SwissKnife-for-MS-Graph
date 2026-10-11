// Package auth acquires Graph tokens via azidentity.
// The token lives only here and in graphapi.Client; it is never exposed outward (ADR-002).
package auth

import (
	"context"
	"errors"
	"fmt"
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

// ErrSessionExpired means the delegated sign-in can no longer be refreshed
// without the user (sign-in frequency policy, password reset, revoked
// session): the operator has to connect again. The message starts with a
// stable code the UI recognises wherever the error surfaces.
var ErrSessionExpired = errors.New("sessionExpired: the sign-in session expired or was revoked — sign in again")

// interactiveCredential signs a user in (device code). After the first
// sign-in it only refreshes silently.
type interactiveCredential interface {
	azcore.TokenCredential
	Authenticate(ctx context.Context, opts *policy.TokenRequestOptions) (azidentity.AuthenticationRecord, error)
}

// TokenProvider implements graphapi.TokenSource on top of azcore.TokenCredential
// with per-resource caching and auto-refresh (proactively, 2 minutes before
// expiry).
type TokenProvider struct {
	cred azcore.TokenCredential
	// interactive is set for delegated sign-in: the first token request
	// prompts (Authenticate); later ones never do — a refresh that needs the
	// user fails with ErrSessionExpired instead of silently starting a new
	// device-code flow nobody sees.
	interactive interactiveCredential
	// signInSem (one slot) serialises the first sign-in only; a channel, not
	// a mutex, so a caller waiting behind it stops when its context ends.
	signInSem chan struct{}
	signedIn  bool // guarded by signInSem

	mu     sync.Mutex // guards tokens; never held across a token request
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
		// Only the explicit first sign-in (Authenticate) may prompt.
		DisableAutomaticAuthentication: true,
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
	return newDelegated(cred), nil
}

func newDelegated(cred interactiveCredential) *TokenProvider {
	return &TokenProvider{cred: cred, interactive: cred, signInSem: make(chan struct{}, 1), tokens: map[string]azcore.AccessToken{}}
}

// Token implements graphapi.TokenSource (Microsoft Graph).
func (p *TokenProvider) Token(ctx context.Context) (string, error) {
	return p.TokenFor(ctx, ResourceGraph)
}

// TokenFor returns a token for the resource (one of the Resource* constants),
// cached per resource. Concurrent refreshes may overlap; the credential has
// its own cache, and no caller waits behind another resource's request.
func (p *TokenProvider) TokenFor(ctx context.Context, resource string) (string, error) {
	p.mu.Lock()
	t, ok := p.tokens[resource]
	p.mu.Unlock()
	if ok && time.Until(t.ExpiresOn) > 2*time.Minute {
		return t.Token, nil
	}

	opts := policy.TokenRequestOptions{Scopes: []string{resource + "/.default"}}
	if err := p.signIn(ctx, opts); err != nil {
		return "", err
	}
	tok, err := p.cred.GetToken(ctx, opts)
	if err != nil {
		return "", mapAuthErr(err)
	}
	p.mu.Lock()
	p.tokens[resource] = tok
	p.mu.Unlock()
	return tok.Token, nil
}

// signIn runs the interactive first sign-in of a delegated provider (a no-op
// for app-only and once signed in). Concurrent first requests wait for it.
func (p *TokenProvider) signIn(ctx context.Context, opts policy.TokenRequestOptions) error {
	if p.interactive == nil {
		return nil
	}
	select {
	case p.signInSem <- struct{}{}:
	case <-ctx.Done():
		return ctx.Err()
	}
	defer func() { <-p.signInSem }()
	if p.signedIn {
		return nil
	}
	if _, err := p.interactive.Authenticate(ctx, &opts); err != nil {
		return err
	}
	p.signedIn = true
	return nil
}

// mapAuthErr turns "the user has to act" into ErrSessionExpired.
func mapAuthErr(err error) error {
	var are *azidentity.AuthenticationRequiredError
	if errors.As(err, &are) {
		return fmt.Errorf("%w (%s)", ErrSessionExpired, err.Error())
	}
	return err
}
