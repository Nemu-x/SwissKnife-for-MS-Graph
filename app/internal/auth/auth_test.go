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
	entered   chan string              // receives the scope of a GetToken that reached its block
	// authStarted/authRelease, when set, hold Authenticate like a user who
	// has not entered the device code yet.
	authStarted chan struct{}
	authRelease chan struct{}
}

func (f *fakeCred) Authenticate(ctx context.Context, _ *policy.TokenRequestOptions) (azidentity.AuthenticationRecord, error) {
	f.mu.Lock()
	f.auths++
	started, release := f.authStarted, f.authRelease
	f.mu.Unlock()
	if release != nil {
		close(started)
		select {
		case <-release:
		case <-ctx.Done():
			return azidentity.AuthenticationRecord{}, ctx.Err()
		}
	}
	return azidentity.AuthenticationRecord{}, nil
}

func (f *fakeCred) GetToken(ctx context.Context, o policy.TokenRequestOptions) (azcore.AccessToken, error) {
	f.mu.Lock()
	err, wait, entered := f.silentErr, f.block[o.Scopes[0]], f.entered
	f.mu.Unlock()
	if wait != nil {
		if entered != nil {
			entered <- o.Scopes[0]
		}
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
	return newDelegated(f)
}

// A caller queued behind a first sign-in that waits for the user gives up
// when its own context ends; the sign-in itself goes on.
func TestWaitForFirstSignInIsCancellable(t *testing.T) {
	f := &fakeCred{authStarted: make(chan struct{}), authRelease: make(chan struct{})}
	p := delegated(f)
	first := make(chan error, 1)
	go func() {
		_, err := p.TokenFor(context.Background(), ResourceGraph)
		first <- err
	}()
	<-f.authStarted

	ctx, cancel := context.WithCancel(context.Background())
	second := make(chan error, 1)
	go func() {
		_, err := p.TokenFor(ctx, ResourceExchange)
		second <- err
	}()
	cancel()
	select {
	case err := <-second:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("err = %v, want context.Canceled", err)
		}
	case <-time.After(2 * time.Second):
		t.Fatal("a cancelled caller must stop waiting for the sign-in")
	}

	close(f.authRelease)
	if err := <-first; err != nil {
		t.Fatal(err)
	}
	if f.auths != 1 {
		t.Fatalf("interactive sign-ins = %d, want 1", f.auths)
	}
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
	f.entered = make(chan string, 1)
	f.mu.Unlock()
	go func() { _, _ = p.TokenFor(context.Background(), ResourceExchange) }()
	<-f.entered // the Exchange request is now blocked inside GetToken

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
