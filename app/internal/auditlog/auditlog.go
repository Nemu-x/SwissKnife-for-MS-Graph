// Package auditlog is a local action journal (ADR-002): what, on what, when, result.
// Format is append-only JSONL in the app data directory, rotated by size:
// actions.log.jsonl is the current file, actions.log.jsonl.1..3 the older ones.
package auditlog

import (
	"bytes"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"sync"
	"time"
)

const (
	// DefaultMaxSize is the size at which the current file is rotated.
	DefaultMaxSize = 10 << 20
	// DefaultKeep is how many rotated files are kept next to the current one.
	DefaultKeep = 3
)

// tailChunk is how much Tail reads per step from the end of a file.
var tailChunk int64 = 64 << 10

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
	mu      sync.Mutex
	path    string
	maxSize int64
	keep    int
}

func New(dir string) *Log {
	return &Log{path: filepath.Join(dir, "actions.log.jsonl"), maxSize: DefaultMaxSize, keep: DefaultKeep}
}

func (l *Log) Path() string { return l.path }

// rotated returns the path of the i-th older file (1 = most recent).
func (l *Log) rotated(i int) string { return l.path + "." + strconv.Itoa(i) }

func (l *Log) Write(e Entry) {
	if e.Time.IsZero() {
		e.Time = time.Now()
	}
	data, err := json.Marshal(e)
	if err != nil {
		return
	}
	data = append(data, '\n')
	l.mu.Lock()
	defer l.mu.Unlock()
	if st, err := os.Stat(l.path); err == nil && st.Size() > 0 && st.Size()+int64(len(data)) > l.maxSize {
		l.rotate()
	}
	f, err := os.OpenFile(l.path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o600)
	if err != nil {
		return
	}
	defer f.Close()
	// Best-effort: a failed audit write must never break the operation itself.
	_, _ = f.Write(data)
}

// rotate shifts the files one step (.2 → .3, .1 → .2, current → .1) and
// drops the oldest. Best-effort: when a rename fails (another process has the
// file open on Windows) the entry is appended to the current file anyway and
// rotation is retried on the next write. Caller holds l.mu.
func (l *Log) rotate() {
	if l.keep <= 0 {
		_ = os.Remove(l.path)
		return
	}
	_ = os.Remove(l.rotated(l.keep))
	for i := l.keep - 1; i >= 1; i-- {
		_ = os.Rename(l.rotated(i), l.rotated(i+1))
	}
	_ = os.Rename(l.path, l.rotated(1))
}

// Tail returns the last n entries (for the Activity tab), oldest first. It
// reads backwards from the end of the current file, continuing into the
// rotated ones only when needed, so its cost does not grow with the log.
// n <= 0 returns everything that is kept.
func (l *Log) Tail(n int) []Entry {
	l.mu.Lock()
	defer l.mu.Unlock()
	files := []string{l.path}
	for i := 1; i <= l.keep; i++ {
		files = append(files, l.rotated(i))
	}
	var out []Entry // newest file first; each part is chronological
	for _, p := range files {
		need := 0
		if n > 0 {
			need = n - len(out)
			if need <= 0 {
				break
			}
		}
		part := tailFile(p, need)
		out = append(part, out...)
	}
	if n > 0 && len(out) > n {
		out = out[len(out)-n:]
	}
	return out
}

// tailFile returns the last need valid entries of one file in chronological
// order (need <= 0: all of them). Malformed lines are skipped.
func tailFile(path string, need int) []Entry {
	f, err := os.Open(path)
	if err != nil {
		return nil
	}
	defer f.Close()
	st, err := f.Stat()
	if err != nil {
		return nil
	}
	pos := st.Size()
	var rev []Entry // newest first
	var carry []byte
	parse := func(line []byte) {
		line = bytes.TrimSpace(line)
		if len(line) == 0 {
			return
		}
		var e Entry
		if json.Unmarshal(line, &e) == nil {
			rev = append(rev, e)
		}
	}
	done := func() bool { return need > 0 && len(rev) >= need }
	for pos > 0 && !done() {
		size := min(tailChunk, pos)
		pos -= size
		buf := make([]byte, size, int(size)+len(carry))
		if _, err := f.ReadAt(buf, pos); err != nil && err != io.EOF {
			return reverse(rev)
		}
		data := append(buf, carry...)
		// Everything after a newline is a complete line: it ends where the
		// already-consumed part (or the file) begins.
		for !done() {
			i := bytes.LastIndexByte(data, '\n')
			if i < 0 {
				break
			}
			parse(data[i+1:])
			data = data[:i]
		}
		carry = append([]byte(nil), data...)
	}
	if pos == 0 && !done() {
		parse(carry) // the first line of the file
	}
	return reverse(rev)
}

func reverse(in []Entry) []Entry {
	for i, j := 0, len(in)-1; i < j; i, j = i+1, j-1 {
		in[i], in[j] = in[j], in[i]
	}
	return in
}
