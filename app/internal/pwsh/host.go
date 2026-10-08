// Package pwsh runs PowerShell 7 for the cmdlets Graph does not cover
// (ADR-008): Exchange Online, Teams, Security & Compliance. One long-lived
// pwsh process per module family keeps the slow Connect-* step to once per
// session. Requests and replies are single JSON lines; the host script runs
// only allow-listed cmdlets and splats their parameters, and tokens travel
// only over stdin — never on the command line or in the environment. (A
// Module Logging policy, if the tenant's machines enable one, still records
// cmdlet parameters — Connect-ExchangeOnline's token included — in the
// PowerShell event log; that is the organization's own audit trail.)
package pwsh

import (
	"bufio"
	"context"
	"crypto/sha256"
	_ "embed"
	"encoding/base64"
	"encoding/binary"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os/exec"
	"strings"
	"sync"
	"sync/atomic"
	"time"
	"unicode/utf16"
)

//go:embed host.ps1
var hostScript []byte

// marker prefixes every reply line; other stdout lines are module noise.
const marker = "\x01SKG "

// Error is a failed cmdlet as reported by the host.
type Error struct {
	Message  string `json:"message"`
	Type     string `json:"type"`
	Category string `json:"category"`
	ErrorID  string `json:"errorId"`
}

func (e *Error) Error() string { return e.Message }

// ErrHostExited means the pwsh process ended (killed by a cancel, crashed).
var ErrHostExited = errors.New("powershell host exited")

type reply struct {
	ID    int64             `json:"id"`
	OK    bool              `json:"ok"`
	Data  []json.RawMessage `json:"data"`
	Error *Error            `json:"error"`
}

// Host is one running pwsh process. Calls are serialized: PowerShell runs one
// pipeline at a time anyway.
type Host struct {
	cmd    *exec.Cmd
	stdin  io.WriteCloser
	lines  chan string
	done   chan struct{}
	stderr *tail

	mu     sync.Mutex
	nextID atomic.Int64
}

// encodedScript is the host script as -EncodedCommand wants it: base64 of
// UTF-16LE. Passing it inline means no script file that a temp cleaner could
// delete, and no file for an AllSigned execution policy to refuse.
var encodedScript = func() string {
	u := utf16.Encode([]rune(string(hostScript)))
	b := make([]byte, 2*len(u))
	for i, c := range u {
		binary.LittleEndian.PutUint16(b[2*i:], c)
	}
	return base64.StdEncoding.EncodeToString(b)
}()

// startTimeout bounds process start plus the init handshake.
const startTimeout = 30 * time.Second

// Start launches pwsh (exe) with the embedded host script and sends the
// cmdlet allow-list.
func Start(exe string, allow []string) (*Host, error) { return StartWithScripts(exe, allow, nil) }

// StartWithScripts also trusts scripts by the SHA-256 of their text (action
// packs); the host refuses any other script.
func StartWithScripts(exe string, allow, scripts []string) (*Host, error) {
	cmd := exec.Command(exe, "-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand", encodedScript)
	cmd.WaitDelay = 2 * time.Second
	hideWindow(cmd)
	stdin, err := cmd.StdinPipe()
	if err != nil {
		return nil, err
	}
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		return nil, err
	}
	h := &Host{cmd: cmd, stdin: stdin, lines: make(chan string, 16), done: make(chan struct{}), stderr: &tail{}}
	cmd.Stderr = h.stderr
	if err := cmd.Start(); err != nil {
		return nil, fmt.Errorf("start powershell: %w", err)
	}
	readDone := make(chan struct{})
	go func() { h.read(stdout); close(readDone) }()
	// Wait closes the pipes, so it runs only once stdout is drained.
	go func() { <-readDone; _ = cmd.Wait(); close(h.done) }()

	ctx, cancel := context.WithTimeout(context.Background(), startTimeout)
	defer cancel()
	if scripts == nil {
		scripts = []string{}
	}
	if _, err := h.call(ctx, map[string]any{"op": "init", "allow": allow, "scripts": scripts}); err != nil {
		h.Close()
		return nil, err
	}
	return h, nil
}

func (h *Host) read(r io.Reader) {
	sc := bufio.NewScanner(r)
	sc.Buffer(make([]byte, 64*1024), 64*1024*1024) // a large result is one line
	for sc.Scan() {
		// Module output written with -NoNewline can precede the marker.
		if line := sc.Text(); strings.Contains(line, marker) {
			h.lines <- line[strings.Index(line, marker)+len(marker):]
		}
	}
	close(h.lines)
}

// call sends one request and waits for its reply. A cancelled context kills
// the process: there is no reliable way to stop a running cmdlet from outside.
func (h *Host) call(ctx context.Context, req map[string]any) ([]json.RawMessage, error) {
	h.mu.Lock()
	defer h.mu.Unlock()
	id := h.nextID.Add(1)
	req["id"] = id
	b, err := json.Marshal(req)
	if err != nil {
		return nil, err
	}
	if _, err := h.stdin.Write(append(b, '\n')); err != nil {
		return nil, h.exitErr()
	}
	for {
		select {
		case <-ctx.Done():
			h.Close()
			return nil, ctx.Err()
		case line, ok := <-h.lines:
			if !ok {
				return nil, h.exitErr()
			}
			var r reply
			if err := json.Unmarshal([]byte(line), &r); err != nil {
				return nil, fmt.Errorf("powershell host: bad reply: %w", err)
			}
			if r.ID != id {
				continue // a reply to a call abandoned earlier
			}
			if !r.OK {
				if r.Error == nil {
					r.Error = &Error{Message: "powershell call failed"}
				}
				return nil, r.Error
			}
			return r.Data, nil
		}
	}
}

func (h *Host) exitErr() error {
	// Give the process a moment to finish so its last stderr lines (e.g. why
	// it refused to start) are in the tail.
	select {
	case <-h.done:
	case <-time.After(time.Second):
	}
	if msg := h.stderr.String(); msg != "" {
		return fmt.Errorf("%w: %s", ErrHostExited, msg)
	}
	return ErrHostExited
}

// Connect signs the family's module in. params carry the token(s) and the
// organization / user principal name.
func (h *Host) Connect(ctx context.Context, family string, params map[string]any) error {
	req := map[string]any{"op": "connect", "family": family}
	for k, v := range params {
		req[k] = v
	}
	_, err := h.call(ctx, req)
	return err
}

// Invoke runs an allow-listed cmdlet; sel limits the returned properties
// (deserialized Exchange objects are large).
func (h *Host) Invoke(ctx context.Context, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error) {
	req := map[string]any{"op": "invoke", "cmdlet": cmdlet, "params": params}
	if len(sel) > 0 {
		req["select"] = sel
	}
	return h.call(ctx, req)
}

// RunScript runs a trusted script with named parameters.
func (h *Host) RunScript(ctx context.Context, script string, params map[string]any) ([]json.RawMessage, error) {
	return h.call(ctx, map[string]any{"op": "script", "script": script, "params": params})
}

// ScriptHash is how the host identifies a trusted script.
func ScriptHash(script string) string {
	sum := sha256.Sum256([]byte(script))
	return hex.EncodeToString(sum[:])
}

// Alive reports whether the process is still running.
func (h *Host) Alive() bool {
	select {
	case <-h.done:
		return false
	default:
		return true
	}
}

// Close ends the process.
func (h *Host) Close() {
	_ = h.stdin.Close()
	if h.cmd.Process != nil && h.Alive() {
		_ = h.cmd.Process.Kill()
	}
}

// tail keeps the last few KB of stderr for error messages.
type tail struct {
	mu  sync.Mutex
	buf []byte
}

func (t *tail) Write(p []byte) (int, error) {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.buf = append(t.buf, p...)
	if len(t.buf) > 4096 {
		t.buf = t.buf[len(t.buf)-4096:]
	}
	return len(p), nil
}

func (t *tail) String() string {
	t.mu.Lock()
	defer t.mu.Unlock()
	return strings.TrimSpace(string(t.buf))
}
