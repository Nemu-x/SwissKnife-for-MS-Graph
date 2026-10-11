package auth

import (
	"context"
	"errors"
	"sync"
	"testing"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore"
	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
)

// fakeCred stands in for a device-code credential: Authenticate is the
// interactive sign-in, GetToken the silent one.
type fakeCred struct {
	mu        sync.Mutex
	auths     int
	silentErr error
	block     map[string]chan struct{} // scope → released when GetToken may return
}

func (f *fakeCred) Authenticate(context.Context, *policy.TokenRequestOptions) (azidentity.AuthenticationRecord, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.auths++
	return azidentity.AuthenticationRecord{}, nil
}

func (f *fakeCred) GetToken(ctx context.Context, o policy.TokenRequestOptions) (azcore.AccessToken, error) {
	f.mu.Lock()
	err, wait := f.silentErr, f.block[o.Scopes[0]]
	f.mu.Unlock()
	if wait != nil {
		select {
		case <-wait:
		case <-ctx.Done():
			return azcore.AccessToken{}, ctx.Err()
		}
	}
	if err != nil {
		return azcore.AccessToken{}, err
	}
	return azcore.AccessToken{Token: "tok:" + o.Scopes[0], ExpiresOn: time.Now().Add(time.Hour)}, nil
}

func delegated(f *fakeCred) *TokenProvider {
	return &TokenProvider{cred: f, interactive: f, tokens: map[string]azcore.AccessToken{}}
}

func TestDelegatedSignsInInteractivelyOnce(t *testing.T) {
	f := &fakeCred{}
	p := delegated(f)
	for _, r := range []string{ResourceGraph, ResourceExchange, ResourceGraph} {
		if _, err := p.TokenFor(context.Background(), r); err != nil {
			t.Fatal(err)
		}
	}
	if f.auths != 1 {
		t.Fatalf("interactive sign-ins = %d, want 1 (only the first connect prompts)", f.auths)
	}
}

func TestRefreshNeedingTheUserIsSessionExpired(t *testing.T) {
	f := &fakeCred{}
	p := delegated(f)
	if _, err := p.TokenFor(context.Background(), ResourceGraph); err != nil {
		t.Fatal(err)
	}
	// Sign-in frequency / revoked session: the silent refresh needs the user.
	f.silentErr = &azidentity.AuthenticationRequiredError{}
	_, err := p.TokenFor(context.Background(), ResourceExchange)
	if !errors.Is(err, ErrSessionExpired) {
		t.Fatalf("err = %v, want ErrSessionExpired", err)
	}
	if f.auths != 1 {
		t.Fatal("an expired session must not start a new device-code flow")
	}
	// Other failures pass through unchanged.
	f.silentErr = errors.New("AADSTS7000215: invalid client secret")
	if _, err := p.TokenFor(context.Background(), ResourceTeams); errors.Is(err, ErrSessionExpired) || err == nil {
		t.Fatalf("a plain failure is no expired session: %v", err)
	}
}

func TestSlowTokenRequestDoesNotBlockOtherResources(t *testing.T) {
	f := &fakeCred{block: map[string]chan struct{}{}}
	p := delegated(f)
	if _, err := p.TokenFor(context.Background(), ResourceGraph); err != nil {
		t.Fatal(err)
	}
	release := make(chan struct{})
	f.mu.Lock()
	f.block[ResourceExchange+"/.default"] = release
	f.mu.Unlock()
	go func() { _, _ = p.TokenFor(context.Background(), ResourceExchange) }()
	time.Sleep(50 * time.Millisecond) // let the Exchange request start

	done := make(chan error, 1)
	go func() {
		_, err := p.TokenFor(context.Background(), ResourceGraph)
		done <- err
	}()
	select {
	case err := <-done:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(2 * time.Second):
		t.Fatal("a cached Graph token must not wait for an Exchange request")
	}
	close(release)
}
