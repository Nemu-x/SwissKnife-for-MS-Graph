// Package corrupt sets aside state files that no longer parse (packs.json,
// clients.json, worker.json …). Loading such a file used to fall back to an
// empty state silently, and the next save then overwrote what was left of
// it. Now the file is renamed to <name>.corrupt-<timestamp> first, and the
// problem is kept until the screen or command that owns the file reports it.
package corrupt

import (
	"fmt"
	"os"
	"path/filepath"
	"sync"
	"time"
)

// Error describes one unreadable file and where it was moved.
type Error struct {
	Path    string // the original file
	MovedTo string // "" when the rename failed
	Err     error  // the parse error
	// RenameErr is why the file could not be set aside (MovedTo == "").
	RenameErr error
}

func (e *Error) Error() string {
	name := filepath.Base(e.Path)
	if e.MovedTo == "" {
		return fmt.Sprintf("%s could not be read (%v) and could not be set aside (%v); the app continues with an empty one — "+
			"copy the file somewhere safe before changing these settings", name, e.Err, e.RenameErr)
	}
	return fmt.Sprintf("%s could not be read (%v): it was set aside as %s and the app continues with an empty one — "+
		"restore it from that copy, or set it up again", name, e.Err, filepath.Base(e.MovedTo))
}

func (e *Error) Unwrap() error { return e.Err }

var (
	mu      sync.Mutex
	pending = map[string]*Error{} // cleaned path → not yet reported
	now     = time.Now
)

// Quarantine renames the unreadable file at path to
// <path>.corrupt-<timestamp>, records the problem for Take and returns it.
func Quarantine(path string, cause error) *Error {
	e := SetAside(path, cause)
	mu.Lock()
	pending[filepath.Clean(path)] = e
	mu.Unlock()
	return e
}

// SetAside is Quarantine for a caller that reports the problem itself:
// nothing is recorded for Take.
func SetAside(path string, cause error) *Error {
	e := &Error{Path: path, Err: cause}
	stamp := now().Format("20060102-150405")
	dst := path + ".corrupt-" + stamp
	// Look for a free name only while the candidate exists; any other Lstat
	// failure (no access to the folder …) ends the search and lets Rename
	// report it.
	for n := 2; n <= 100; n++ {
		if _, err := os.Lstat(dst); err != nil {
			break
		}
		dst = fmt.Sprintf("%s.corrupt-%s-%d", path, stamp, n)
	}
	if err := os.Rename(path, dst); err != nil {
		e.RenameErr = err
	} else {
		e.MovedTo = dst
	}
	return e
}

// Take returns the unreported problem with path, if any, and forgets it:
// each quarantine is reported once.
func Take(path string) error {
	mu.Lock()
	defer mu.Unlock()
	key := filepath.Clean(path)
	e, ok := pending[key]
	if !ok {
		return nil
	}
	delete(pending, key)
	return e
}

// Peek returns the unreported problem with path without forgetting it, for
// callers that only need to explain why something is unavailable.
func Peek(path string) error {
	mu.Lock()
	defer mu.Unlock()
	if e, ok := pending[filepath.Clean(path)]; ok {
		return e
	}
	return nil
}
