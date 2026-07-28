package services

import (
	"errors"
	"net/http"
	"strings"
	"sync"
	"testing"
	"time"

	"swissknife-app/internal/ops"
)

// captureEvents routes emitOp/emitEvent into a buffer for the test's duration.
func captureEvents(t *testing.T) *eventBuffer {
	t.Helper()
	buf := &eventBuffer{}
	eventSink = buf.add
	t.Cleanup(func() { eventSink = nil })
	return buf
}

type eventBuffer struct {
	mu     sync.Mutex
	events []capturedEvent
}

type capturedEvent struct {
	name string
	data map[string]any
}

func (b *eventBuffer) add(name string, data map[string]any) {
	b.mu.Lock()
	defer b.mu.Unlock()
	// Copy: emitters may reuse maps.
	cp := make(map[string]any, len(data))
	for k, v := range data {
		cp[k] = v
	}
	b.events = append(b.events, capturedEvent{name: name, data: cp})
}

func (b *eventBuffer) snapshot() []capturedEvent {
	b.mu.Lock()
	defer b.mu.Unlock()
	return append([]capturedEvent(nil), b.events...)
}

// TestCancelAbortsInFlightHTTP: cancelling the transfer op must abort the HTTP
// request that is currently blocked on the server, not wait it out.
func TestCancelAbortsInFlightHTTP(t *testing.T) {
	release := make(chan struct{})
	reached := make(chan struct{}, 1)
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/users/arch@contoso.com/drive") {
			select {
			case reached <- struct{}{}: // deterministic: the copy IS in flight now
			default:
			}
			<-release // block the server-side-copy setup call until the test ends
		}
		w.Write([]byte(`{}`))
	})
	t.Cleanup(func() { close(release) })
	drive := NewDriveService(sess)

	done := make(chan struct{})
	var res *CopyResult
	go func() {
		defer close(done)
		res, _ = drive.CopyBetweenUsers("alice@contoso.com", "arch@contoso.com", "backup", false)
	}()

	// Wait until the request is provably blocked on the server, then cancel.
	select {
	case <-reached:
	case <-time.After(3 * time.Second):
		t.Fatal("copy never reached the blocked call")
	}
	start := time.Now()
	sess.Ops.CancelKind(ops.KindTransfer)
	select {
	case <-done:
	case <-time.After(3 * time.Second):
		t.Fatal("copy did not return after cancel — in-flight HTTP was not aborted")
	}
	if elapsed := time.Since(start); elapsed > 2*time.Second {
		t.Fatalf("cancel took %v — should abort the blocked request immediately", elapsed)
	}
	if res == nil || !res.Canceled {
		t.Fatalf("result must be marked canceled, got %+v", res)
	}
}

// TestTransferSingleFlight: a second copy while one is live must be rejected
// immediately with a typed error.
func TestTransferSingleFlight(t *testing.T) {
	release := make(chan struct{})
	reached := make(chan struct{}, 1)
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "GET" && strings.HasSuffix(r.URL.Path, "/users/arch@contoso.com/drive") {
			select {
			case reached <- struct{}{}:
			default:
			}
			<-release
		}
		w.Write([]byte(`{}`))
	})
	drive := NewDriveService(sess)

	done := make(chan struct{})
	go func() {
		defer close(done)
		_, _ = drive.CopyBetweenUsers("alice@contoso.com", "arch@contoso.com", "backup", false)
	}()
	select {
	case <-reached: // first copy provably holds the transfer slot
	case <-time.After(3 * time.Second):
		t.Fatal("first copy never reached the blocked call")
	}

	_, err := drive.CopyBetweenUsers("bob@contoso.com", "arch@contoso.com", "backup2", false)
	var are *ops.AlreadyRunningError
	if err == nil || !errors.As(err, &are) {
		t.Fatalf("second copy must fail with AlreadyRunningError, got %v", err)
	}
	close(release)
	<-done
}

// TestPlaybookAndChildCopyEmitDistinctOpIds: a playbook run with a backup must
// stamp playbook events with the playbook opId and transfer events with the
// child copy's own opId — never mixed, never missing.
func TestPlaybookAndChildCopyEmitDistinctOpIds(t *testing.T) {
	buf := captureEvents(t)
	var calls []string
	sess := harness(t, offboardHarness(t, &calls, map[string]string{}))
	pb := NewPlaybookService(sess)

	req := fullOffboardRequest()
	req.BackupToUser = "archive@contoso.com"
	if _, err := pb.Offboard(req); err != nil {
		t.Fatal(err)
	}

	playbookIds := map[string]bool{}
	transferIds := map[string]bool{}
	for _, ev := range buf.snapshot() {
		id, _ := ev.data["opId"].(string)
		kind, _ := ev.data["opKind"].(string)
		if id == "" || kind == "" {
			t.Fatalf("event %s missing op envelope: %+v", ev.name, ev.data)
		}
		switch {
		case strings.HasPrefix(ev.name, "playbook:"):
			if kind != "playbook" {
				t.Fatalf("playbook event stamped %q", kind)
			}
			playbookIds[id] = true
		case strings.HasPrefix(ev.name, "transfer:"):
			if kind != "transfer" {
				t.Fatalf("transfer event stamped %q", kind)
			}
			transferIds[id] = true
		}
	}
	if len(playbookIds) != 1 || len(transferIds) != 1 {
		t.Fatalf("want exactly one opId per kind, got playbook=%v transfer=%v", playbookIds, transferIds)
	}
	for id := range playbookIds {
		if transferIds[id] {
			t.Fatal("playbook and child transfer must not share an opId")
		}
	}
}
