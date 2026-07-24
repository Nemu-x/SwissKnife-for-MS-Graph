// Package journal persists every long-running operation to disk as one
// append-only JSONL file per run (runs/<opId>.jsonl): a begin record, step /
// item / log events as they happen, and a terminal end record. Each line is
// flushed on write, so a crash loses at most the final line — a run without an
// end record renders as "interrupted" and (for cloud copies) can be resumed.
package journal

import (
	"bufio"
	"encoding/json"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"
)

// Entry is one JSONL line. Type is "begin" | "step" | "item" | "log" | "end".
type Entry struct {
	Type string         `json:"t"`
	At   time.Time      `json:"at"`
	Data map[string]any `json:"data,omitempty"`
}

// RunSummary is the cheap listing view: the begin record plus, when present,
// the end record (read from the first and last lines only).
type RunSummary struct {
	OpID        string         `json:"opId"`
	Kind        string         `json:"kind"`
	Target      string         `json:"target"`
	StartedAt   time.Time      `json:"startedAt"`
	EndedAt     *time.Time     `json:"endedAt,omitempty"` // nil = interrupted or still running
	Summary     map[string]any `json:"summary,omitempty"`
	Interrupted bool           `json:"interrupted"`
}

// Run is a fully parsed journal file.
type Run struct {
	RunSummary
	Begin  map[string]any `json:"begin"`
	Events []Entry       `json:"events"`
}

// Log owns the runs directory. Writers append via Begin/Event/End; files stay
// open between writes and close on End (or Close).
type Log struct {
	dir  string
	mu   sync.Mutex
	open map[string]*os.File
	live map[string]bool // opIds currently writing (not "interrupted")
}

func New(dir string) *Log {
	_ = os.MkdirAll(dir, 0o755)
	return &Log{dir: dir, open: map[string]*os.File{}, live: map[string]bool{}}
}

func (l *Log) path(opID string) string {
	// opIds are generated (hex + dash) but never trust them as path segments.
	safe := strings.Map(func(r rune) rune {
		switch {
		case r >= 'a' && r <= 'z', r >= '0' && r <= '9', r == '-':
			return r
		}
		return '_'
	}, opID)
	return filepath.Join(l.dir, safe+".jsonl")
}

func (l *Log) write(opID string, rec Entry) {
	l.mu.Lock()
	defer l.mu.Unlock()
	f, ok := l.open[opID]
	if !ok {
		var err error
		f, err = os.OpenFile(l.path(opID), os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
		if err != nil {
			return // journaling is best-effort; never fail the operation itself
		}
		l.open[opID] = f
	}
	b, err := json.Marshal(rec)
	if err != nil {
		return
	}
	_, _ = f.Write(append(b, '\n'))
	_ = f.Sync()
}

// Begin starts a run's journal. data should carry kind/target plus anything
// needed to resume (e.g. the copy plan).
func (l *Log) Begin(opID string, data map[string]any) {
	l.mu.Lock()
	l.live[opID] = true
	l.mu.Unlock()
	l.write(opID, Entry{Type: "begin", At: time.Now(), Data: data})
}

// Event appends a step/item/log record.
func (l *Log) Event(opID, typ string, data map[string]any) {
	l.write(opID, Entry{Type: typ, At: time.Now(), Data: data})
}

// End writes the terminal record and closes the file.
func (l *Log) End(opID string, summary map[string]any) {
	l.write(opID, Entry{Type: "end", At: time.Now(), Data: summary})
	l.mu.Lock()
	defer l.mu.Unlock()
	if f, ok := l.open[opID]; ok {
		_ = f.Close()
		delete(l.open, opID)
	}
	delete(l.live, opID)
}

// Close closes all open journal files (app shutdown, tests). Runs without an
// end record will list as interrupted afterwards.
func (l *Log) Close() {
	l.mu.Lock()
	defer l.mu.Unlock()
	for id, f := range l.open {
		_ = f.Close()
		delete(l.open, id)
	}
	l.live = map[string]bool{}
}

// List returns the newest limit runs (opIds are time-sortable, so the file
// name order is the time order). Each file is scanned once to find its first
// and last lines — full events are parsed only by Get.
func (l *Log) List(limit int) []RunSummary {
	entries, err := os.ReadDir(l.dir)
	if err != nil {
		return nil
	}
	names := make([]string, 0, len(entries))
	for _, e := range entries {
		if !e.IsDir() && strings.HasSuffix(e.Name(), ".jsonl") {
			names = append(names, e.Name())
		}
	}
	sort.Sort(sort.Reverse(sort.StringSlice(names)))
	if limit > 0 && len(names) > limit {
		names = names[:limit]
	}
	l.mu.Lock()
	liveNow := make(map[string]bool, len(l.live))
	for k := range l.live {
		liveNow[k] = true
	}
	l.mu.Unlock()
	out := make([]RunSummary, 0, len(names))
	for _, name := range names {
		opID := strings.TrimSuffix(name, ".jsonl")
		first, last, err := firstLastLines(filepath.Join(l.dir, name))
		if err != nil {
			continue
		}
		var begin Entry
		if json.Unmarshal([]byte(first), &begin) != nil || begin.Type != "begin" {
			continue
		}
		s := RunSummary{OpID: opID, StartedAt: begin.At}
		s.Kind, _ = begin.Data["kind"].(string)
		s.Target, _ = begin.Data["target"].(string)
		var end Entry
		if json.Unmarshal([]byte(last), &end) == nil && end.Type == "end" {
			at := end.At
			s.EndedAt = &at
			s.Summary = end.Data
		} else {
			s.Interrupted = !liveNow[opID]
		}
		out = append(out, s)
	}
	return out
}

// Get parses a full run journal.
func (l *Log) Get(opID string) (*Run, error) {
	f, err := os.Open(l.path(opID))
	if err != nil {
		return nil, err
	}
	defer f.Close()
	run := &Run{}
	sc := bufio.NewScanner(f)
	sc.Buffer(make([]byte, 0, 64*1024), 4<<20)
	for sc.Scan() {
		var rec Entry
		if json.Unmarshal(sc.Bytes(), &rec) != nil {
			continue // torn final line after a crash
		}
		switch rec.Type {
		case "begin":
			run.Begin = rec.Data
			run.OpID = opID
			run.StartedAt = rec.At
			run.Kind, _ = rec.Data["kind"].(string)
			run.Target, _ = rec.Data["target"].(string)
		case "end":
			at := rec.At
			run.EndedAt = &at
			run.Summary = rec.Data
		default:
			run.Events = append(run.Events, rec)
		}
	}
	// An oversized line (ErrTooLong) would silently truncate the events the
	// resume logic depends on — surface it instead of returning a partial run.
	if err := sc.Err(); err != nil {
		return nil, err
	}
	if run.Begin == nil {
		return nil, os.ErrNotExist
	}
	l.mu.Lock()
	live := l.live[opID]
	l.mu.Unlock()
	run.Interrupted = run.EndedAt == nil && !live
	return run, nil
}

// Prune keeps the newest keep runs and deletes the rest (called at startup).
// Live/open runs are never candidates, so a future mid-session caller cannot
// delete a journal that is still being written.
func (l *Log) Prune(keep int) {
	entries, err := os.ReadDir(l.dir)
	if err != nil {
		return
	}
	l.mu.Lock()
	active := make(map[string]bool, len(l.open)+len(l.live))
	for id := range l.open {
		active[id] = true
	}
	for id := range l.live {
		active[id] = true
	}
	l.mu.Unlock()
	names := make([]string, 0, len(entries))
	for _, e := range entries {
		if !e.IsDir() && strings.HasSuffix(e.Name(), ".jsonl") && !active[strings.TrimSuffix(e.Name(), ".jsonl")] {
			names = append(names, e.Name())
		}
	}
	if len(names) <= keep {
		return
	}
	sort.Strings(names) // oldest first
	for _, name := range names[:len(names)-keep] {
		_ = os.Remove(filepath.Join(l.dir, name))
	}
}

// firstLastLines reads a file's first and last non-empty lines cheaply.
func firstLastLines(path string) (string, string, error) {
	f, err := os.Open(path)
	if err != nil {
		return "", "", err
	}
	defer f.Close()
	sc := bufio.NewScanner(f)
	sc.Buffer(make([]byte, 0, 64*1024), 4<<20)
	first, last := "", ""
	for sc.Scan() {
		line := strings.TrimSpace(sc.Text())
		if line == "" {
			continue
		}
		if first == "" {
			first = line
		}
		last = line
	}
	if first == "" {
		return "", "", os.ErrNotExist
	}
	return first, last, nil
}

