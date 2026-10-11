package pwsh

import (
	"sync/atomic"
	"testing"
	"time"
)

// A slow host start for one key must not hold up the others; callers of the
// same key wait for that one start instead of starting a second host.
func TestPoolSlowStartBlocksOnlyItsKey(t *testing.T) {
	p := fakePool(t, "Get-Connects")
	env := poolEnv(t)
	inner := p.start
	release := make(chan struct{})
	var packStarts atomic.Int32
	p.start = func(exe string, allow []string, scripts []TrustedScript) (*Host, error) {
		if allow == nil { // the pack host
			packStarts.Add(1)
			<-release
			allow = []string{} // the fake host wants a list
		}
		return inner(exe, allow, scripts)
	}
	type res struct {
		h   *Host
		err error
	}
	slow := make(chan res, 2)
	for range 2 {
		go func() {
			h, err := p.host(env, FamilyExchange+packSuffix)
			slow <- res{h, err}
		}()
	}
	// Wait until the first pack start is in flight.
	deadline := time.Now().Add(5 * time.Second)
	for packStarts.Load() == 0 && time.Now().Before(deadline) {
		time.Sleep(10 * time.Millisecond)
	}

	fast := make(chan error, 1)
	go func() {
		_, err := p.Invoke(env, FamilyExchange, "Get-Connects", nil)
		fast <- err
	}()
	select {
	case err := <-fast:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(10 * time.Second):
		t.Fatal("another key must not wait for the slow start")
	}
	close(release)
	a, b := <-slow, <-slow
	if a.err != nil || b.err != nil {
		t.Fatalf("pack callers: %v / %v", a.err, b.err)
	}
	if n := packStarts.Load(); n != 1 || a.h != b.h {
		t.Fatalf("pack host started %d times, want 1 shared (the second caller waits)", n)
	}
}

// A pack host whose start straddles ResetPacks runs the old scripts: it is
// replaced at its next use. ResetPacks does not wait for the start.
func TestPoolPackHostStartedAcrossResetIsStale(t *testing.T) {
	p := fakePool(t, "Get-Connects")
	env := poolEnv(t)
	inner := p.start
	started := make(chan struct{})
	release := make(chan struct{})
	p.start = func(exe string, _ []string, scripts []TrustedScript) (*Host, error) {
		close(started)
		<-release
		return inner(exe, []string{}, scripts)
	}
	errc := make(chan error, 1)
	go func() {
		_, err := p.host(env, FamilyExchange+packSuffix)
		errc <- err
	}()
	<-started
	done := make(chan struct{})
	go func() { p.ResetPacks(); close(done) }()
	select {
	case <-done:
	case <-time.After(2 * time.Second):
		t.Fatal("ResetPacks must not wait for a start in flight")
	}
	close(release)
	if err := <-errc; err != nil {
		t.Fatal(err)
	}
	if e := p.hosts[FamilyExchange+packSuffix]; e == nil || !e.stale {
		t.Fatal("a pack host started before ResetPacks must be stale")
	}
}

// A CloseFor that lands while a host is starting discards that host.
func TestPoolCloseForDuringStartDiscardsHost(t *testing.T) {
	p := fakePool(t, "Get-Connects")
	env := poolEnv(t)
	inner := p.start
	started := make(chan struct{})
	release := make(chan struct{})
	p.start = func(exe string, allow []string, scripts []TrustedScript) (*Host, error) {
		close(started)
		<-release
		return inner(exe, allow, scripts)
	}
	errc := make(chan error, 1)
	go func() {
		_, err := p.Invoke(env, FamilyExchange, "Get-Connects", nil)
		errc <- err
	}()
	<-started
	closed := make(chan struct{})
	go func() { p.CloseFor(env.Graph); close(closed) }()
	select {
	case <-closed:
	case <-time.After(2 * time.Second):
		t.Fatal("CloseFor must not wait for a start in flight")
	}
	close(release)
	if err := <-errc; err == nil {
		t.Fatal("a host started across CloseFor must not be used")
	}
	if len(p.hosts) != 0 {
		t.Fatal("nor kept")
	}
}
