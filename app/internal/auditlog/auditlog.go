// Package auditlog is a local action journal (ADR-002): what, on what, when, result.
// Format is append-only JSONL in the app data directory.
package auditlog

import (
	"encoding/json"
	"os"
	"path/filepath"
	"sync"
	"time"
)

type Entry struct {
	Time    time.Time `json:"time"`
	Action  string    `json:"action"`           // e.g. "intune.wipe"
	Target  string    `json:"target"`           // UPN / device id / item id
	Detail  string    `json:"detail,omitempty"` // short params, WITHOUT anything sensitive
	OK      bool      `json:"ok"`
	Error   string    `json:"error,omitempty"`
	Profile string    `json:"profile,omitempty"` // connection profile name
}

type Log struct {
	mu   sync.Mutex
	path string
}

func New(dir string) *Log {
	return &Log{path: filepath.Join(dir, "actions.log.jsonl")}
}

func (l *Log) Path() string { return l.path }

func (l *Log) Write(e Entry) {
	if e.Time.IsZero() {
		e.Time = time.Now()
	}
	data, err := json.Marshal(e)
	if err != nil {
		return
	}
	l.mu.Lock()
	defer l.mu.Unlock()
	f, err := os.OpenFile(l.path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o600)
	if err != nil {
		return
	}
	defer f.Close()
	// Best-effort: a failed audit write must never break the operation itself.
	_, _ = f.Write(append(data, '\n'))
}

// Tail returns the last n entries (for the Activity tab).
func (l *Log) Tail(n int) []Entry {
	l.mu.Lock()
	defer l.mu.Unlock()
	data, err := os.ReadFile(l.path)
	if err != nil {
		return nil
	}
	var out []Entry
	start := 0
	for i := 0; i <= len(data); i++ {
		if i == len(data) || data[i] == '\n' {
			if i > start {
				var e Entry
				if json.Unmarshal(data[start:i], &e) == nil {
					out = append(out, e)
				}
			}
			start = i + 1
		}
	}
	if n > 0 && len(out) > n {
		out = out[len(out)-n:]
	}
	return out
}
