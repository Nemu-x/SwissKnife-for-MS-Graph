package auditlog

import (
	"fmt"
	"os"
	"strings"
	"sync"
	"testing"
)

func targets(es []Entry) string {
	var s []string
	for _, e := range es {
		s = append(s, e.Target)
	}
	return strings.Join(s, ",")
}

func TestTailReturnsLastEntriesInOrder(t *testing.T) {
	l := New(t.TempDir())
	for i := 1; i <= 5; i++ {
		l.Write(Entry{Action: "a", Target: fmt.Sprint(i), OK: true})
	}
	if got := targets(l.Tail(3)); got != "3,4,5" {
		t.Errorf("Tail(3) = %s", got)
	}
	if got := targets(l.Tail(0)); got != "1,2,3,4,5" {
		t.Errorf("Tail(0) = %s", got)
	}
	if got := targets(l.Tail(50)); got != "1,2,3,4,5" {
		t.Errorf("Tail(50) = %s", got)
	}
}

func TestTailMissingFile(t *testing.T) {
	if got := New(t.TempDir()).Tail(10); len(got) != 0 {
		t.Errorf("Tail on a missing log = %v", got)
	}
}

func TestRotationAtThresholdKeepsThreeFiles(t *testing.T) {
	l := New(t.TempDir())
	one := Entry{Action: "a", Target: "000", OK: true}
	l.Write(one)
	st, _ := os.Stat(l.Path())
	lineSize := st.Size()
	_ = os.Remove(l.Path())
	l.maxSize = lineSize * 3 // three entries per file

	for i := 0; i < 20; i++ {
		l.Write(Entry{Action: "a", Target: fmt.Sprintf("%03d", i), OK: true})
	}
	for _, p := range []string{l.Path(), l.rotated(1), l.rotated(2), l.rotated(3)} {
		st, err := os.Stat(p)
		if err != nil {
			t.Fatalf("%s missing: %v", p, err)
		}
		if st.Size() > l.maxSize {
			t.Errorf("%s is %d bytes, over the %d cap", p, st.Size(), l.maxSize)
		}
	}
	if _, err := os.Stat(l.rotated(4)); err == nil {
		t.Error("a fourth rotated file was kept")
	}
	// 20 writes, 3 per file, 4 files: the last 11 survive (2 in the current).
	if got := targets(l.Tail(0)); got != "009,010,011,012,013,014,015,016,017,018,019" {
		t.Errorf("Tail(0) = %s", got)
	}
	// Tail crosses from the current file into the rotated ones.
	if got := targets(l.Tail(5)); got != "015,016,017,018,019" {
		t.Errorf("Tail(5) across rotation = %s", got)
	}
}

func TestTailLargeFileReadsBackwardsAcrossChunks(t *testing.T) {
	old := tailChunk
	tailChunk = 100 // lines straddle chunk boundaries
	t.Cleanup(func() { tailChunk = old })

	dir := t.TempDir()
	l := New(dir)
	var b strings.Builder
	for i := 0; i < 20000; i++ {
		fmt.Fprintf(&b, `{"time":"2026-01-01T00:00:00Z","action":"a","target":"%05d","ok":true}`+"\n", i)
		if i == 19997 {
			b.WriteString("not json\n\n") // skipped, does not count
		}
	}
	if err := os.WriteFile(l.Path(), []byte(b.String()), 0o600); err != nil {
		t.Fatal(err)
	}
	if got := targets(l.Tail(4)); got != "19996,19997,19998,19999" {
		t.Errorf("Tail(4) = %s", got)
	}
	// A file without a trailing newline still yields its last line.
	if err := os.WriteFile(l.Path(), []byte(`{"target":"x"}`+"\n"+`{"target":"y"}`), 0o600); err != nil {
		t.Fatal(err)
	}
	if got := targets(l.Tail(5)); got != "x,y" {
		t.Errorf("Tail without trailing newline = %s", got)
	}
}

func TestConcurrentWritesStayWholeLines(t *testing.T) {
	l := New(t.TempDir())
	l.maxSize = 2048
	var wg sync.WaitGroup
	for g := 0; g < 8; g++ {
		wg.Add(1)
		go func(g int) {
			defer wg.Done()
			for i := 0; i < 50; i++ {
				l.Write(Entry{Action: "a", Target: fmt.Sprintf("%d-%d", g, i), OK: true})
			}
		}(g)
	}
	wg.Wait()
	got := l.Tail(0)
	if len(got) == 0 {
		t.Fatal("nothing kept")
	}
	for _, e := range got {
		if e.Action != "a" {
			t.Fatalf("garbled entry %+v", e)
		}
	}
}
