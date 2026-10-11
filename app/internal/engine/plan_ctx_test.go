package engine

import (
	"context"
	"errors"
	"testing"
	"time"
)

// slowPlanner blocks its preview until the context ends.
type slowPlanner struct{ started chan struct{} }

func (slowPlanner) Backend() Backend { return BackendGraph }
func (p slowPlanner) Plan(env Env, _ Inputs) ([]Change, error) {
	close(p.started)
	<-env.Ctx.Done()
	return nil, env.Ctx.Err()
}
func (slowPlanner) Apply(Env, Inputs, Change) error { return nil }

func slowEngine(t *testing.T) (*Engine, chan struct{}) {
	t.Helper()
	e := New(newSession(t), fakeProvider{backend: BackendGraph})
	started := make(chan struct{})
	e.Register(Action{Manifest: Manifest{ID: "slow", Danger: Write}, Impls: []Impl{slowPlanner{started}}})
	return e, started
}

func TestPlanCancelStopsThePreview(t *testing.T) {
	e, started := slowEngine(t)
	ctx, cancel := context.WithCancel(context.Background())
	errc := make(chan error, 1)
	go func() {
		_, err := e.PlanCtx(ctx, "slow", nil)
		errc <- err
	}()
	<-started
	cancel()
	select {
	case err := <-errc:
		var ee *Error
		if !errors.As(err, &ee) || ee.Code != "planCanceled" {
			t.Fatalf("err = %v, want planCanceled", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("a cancelled preview must return")
	}
}

func TestPlanTimeoutIsReported(t *testing.T) {
	e, _ := slowEngine(t)
	ctx, cancel := context.WithTimeout(context.Background(), 50*time.Millisecond)
	defer cancel()
	_, err := e.PlanCtx(ctx, "slow", nil)
	var ee *Error
	if !errors.As(err, &ee) || ee.Code != "planTimeout" {
		t.Fatalf("err = %v, want planTimeout", err)
	}
}
