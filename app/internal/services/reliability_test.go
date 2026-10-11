package services

import (
	"context"
	"errors"
	"net/http"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/secrets"
)

func TestConnectRefusedWhileAWriteRuns(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	c := NewConnectService(sess, secrets.NewStoreAt(t.TempDir()))
	op, err := sess.Ops.Start(context.Background(), ops.KindBulk)
	if err != nil {
		t.Fatal(err)
	}
	_, err = c.Connect(ConnectRequest{ProfileID: "other"})
	if err == nil || !strings.Contains(err.Error(), `"code":"connectBusy"`) {
		t.Fatalf("err = %v, want connectBusy", err)
	}
	if !sess.Connected() {
		t.Fatal("a refused switch must leave the connection alone")
	}
	sess.Ops.Finish(op)
	// Nothing running: the switch goes ahead (and fails on the unknown profile).
	if _, err := c.Connect(ConnectRequest{ProfileID: "other"}); err == nil || strings.Contains(err.Error(), "connectBusy") {
		t.Fatalf("err = %v, want the profile lookup error", err)
	}
	// A read fan-out does not block a switch; it is cancelled by it.
	fan, _ := sess.Ops.Start(context.Background(), ops.KindFanOut)
	defer sess.Ops.Finish(fan)
	if err := c.refuseWhileBusy(); err != nil {
		t.Fatalf("fan-out must not block: %v", err)
	}
}

// slowPreview blocks its preview until the context ends.
type slowPreview struct{ started chan struct{} }

func (slowPreview) Backend() engine.Backend { return engine.BackendGraph }
func (p slowPreview) Plan(env engine.Env, _ engine.Inputs) ([]engine.Change, error) {
	close(p.started)
	<-env.Ctx.Done()
	return nil, env.Ctx.Err()
}
func (slowPreview) Apply(engine.Env, engine.Inputs, engine.Change) error { return nil }

func startSlowPreview(t *testing.T) (*ActionsService, chan error) {
	t.Helper()
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	a := NewActionsService(sess, secrets.NewStoreAt(t.TempDir()))
	started := make(chan struct{})
	a.e.Register(engine.Action{Manifest: engine.Manifest{ID: "test.slow", Danger: engine.Write}, Impls: []engine.Impl{slowPreview{started}}})
	errc := make(chan error, 1)
	go func() {
		_, err := a.Plan("test.slow", nil, "p-1")
		errc <- err
	}()
	select {
	case <-started:
	case <-time.After(5 * time.Second):
		t.Fatal("preview did not start")
	}
	return a, errc
}

func waitErr(t *testing.T, errc chan error, code string) {
	t.Helper()
	select {
	case err := <-errc:
		if err == nil || !strings.Contains(err.Error(), `"code":"`+code+`"`) {
			t.Fatalf("err = %v, want %s", err, code)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("the preview must stop")
	}
}

func TestCancelStopsARunningPreview(t *testing.T) {
	a, errc := startSlowPreview(t)
	a.Cancel("p-1")
	waitErr(t, errc, "planCanceled")
}

func TestDisconnectStopsARunningPreview(t *testing.T) {
	a, errc := startSlowPreview(t)
	a.s.Disconnect()
	waitErr(t, errc, "planCanceled")
}

func TestSessionExpiredGetsItsOwnCode(t *testing.T) {
	err := wrapOpErr(errors.Join(errors.New("graph: acquire token"), auth.ErrSessionExpired))
	var oe *OpError
	if !errors.As(err, &oe) || oe.Code != "sessionExpired" {
		t.Fatalf("err = %v, want the sessionExpired envelope", err)
	}
}
