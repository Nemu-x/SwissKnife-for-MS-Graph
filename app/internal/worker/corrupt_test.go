package worker

import (
	"context"
	"fmt"
	"net"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"swissknife-app/internal/auditlog"
)

func TestCorruptClientsFileIsSetAsideAndLogged(t *testing.T) {
	dir := t.TempDir()
	id, err := LoadOrCreateIdentity(dir, "worker")
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "worker-clients.json")
	if err := os.WriteFile(path, []byte("[{"), 0o600); err != nil {
		t.Fatal(err)
	}
	var logged []string
	s := &Server{Identity: id, Dir: dir, Audit: auditlog.New(dir),
		Logf: func(f string, a ...any) { logged = append(logged, fmt.Sprintf(f, a...)) }}
	if got := s.Clients(); len(got) != 0 {
		t.Fatalf("clients = %v", got)
	}
	if len(logged) != 1 || !strings.Contains(logged[0], "worker-clients.json.corrupt-") {
		t.Fatalf("log = %v", logged)
	}
	if m, _ := filepath.Glob(path + ".corrupt-*"); len(m) != 1 {
		t.Fatalf("set-aside copies = %v", m)
	}
	// Reported once, not on every call.
	_ = s.Clients()
	if len(logged) != 1 {
		t.Errorf("logged again: %v", logged)
	}
	if b, _ := os.ReadFile(filepath.Join(dir, "actions.log.jsonl")); !strings.Contains(string(b), "worker.clientsCorrupt") {
		t.Errorf("no audit entry: %s", b)
	}
}

func TestServeReportsACorruptClientsFileAtStart(t *testing.T) {
	dir := t.TempDir()
	id, err := LoadOrCreateIdentity(dir, "worker")
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "worker-clients.json"), []byte("{"), 0o600); err != nil {
		t.Fatal(err)
	}
	logged := make(chan string, 4)
	s := &Server{Identity: id, Dir: dir, Audit: auditlog.New(dir),
		Logf: func(f string, a ...any) { logged <- fmt.Sprintf(f, a...) }}
	l, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	go func() { _ = s.ServeListener(ctx, l) }()
	// No client call: the warning comes from starting to serve.
	select {
	case msg := <-logged:
		if !strings.Contains(msg, "worker-clients.json") {
			t.Fatalf("log = %q", msg)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("no warning at start")
	}
	if b, _ := os.ReadFile(filepath.Join(dir, "actions.log.jsonl")); !strings.Contains(string(b), "worker.clientsCorrupt") {
		t.Errorf("no audit entry: %s", b)
	}
}
