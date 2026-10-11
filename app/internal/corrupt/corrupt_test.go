package corrupt

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestQuarantineMovesFileAndReportsOnce(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "packs.json")
	if err := os.WriteFile(path, []byte("{broken"), 0o600); err != nil {
		t.Fatal(err)
	}
	e := Quarantine(path, errors.New("unexpected end of JSON input"))
	if e.MovedTo == "" || !strings.HasPrefix(filepath.Base(e.MovedTo), "packs.json.corrupt-") {
		t.Fatalf("MovedTo = %q", e.MovedTo)
	}
	if _, err := os.Stat(path); !os.IsNotExist(err) {
		t.Error("the corrupt file is still in place")
	}
	if b, _ := os.ReadFile(e.MovedTo); string(b) != "{broken" {
		t.Errorf("set-aside copy = %q", b)
	}
	err := Take(path)
	var ce *Error
	if !errors.As(err, &ce) || !strings.Contains(err.Error(), filepath.Base(e.MovedTo)) {
		t.Fatalf("Take = %v", err)
	}
	if Take(path) != nil {
		t.Error("a problem was reported twice")
	}

	// A second quarantine in the same second does not clobber the first copy.
	_ = os.WriteFile(path, []byte("{again"), 0o600)
	e2 := Quarantine(path, errors.New("bad"))
	if e2.MovedTo == "" || e2.MovedTo == e.MovedTo {
		t.Errorf("second copy = %q (first %q)", e2.MovedTo, e.MovedTo)
	}
	_ = Take(path)
}
