package graphapi

import (
	"context"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync/atomic"
	"testing"
	"time"
)

// TestChunkUploadRetriesOn429: a throttled chunk must be retried (honoring
// Retry-After) and the upload must complete instead of failing the file.
func TestChunkUploadRetriesOn429(t *testing.T) {
	var chunkPuts, throttled atomic.Int32
	var srvURL string
	c, srv := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "POST" && strings.HasSuffix(r.URL.Path, "createUploadSession"):
			fmt.Fprintf(w, `{"uploadUrl":"%s/upload-session"}`, srvURL)
		case r.Method == "PUT" && r.URL.Path == "/upload-session":
			n := chunkPuts.Add(1)
			if n == 1 { // throttle the very first chunk attempt
				throttled.Add(1)
				w.Header().Set("Retry-After", "1")
				w.WriteHeader(429)
				fmt.Fprint(w, `{"error":{"code":"activityLimitReached","message":"throttled"}}`)
				return
			}
			if strings.HasPrefix(r.Header.Get("Content-Range"), "bytes 0-") {
				w.WriteHeader(202) // first chunk accepted
				return
			}
			w.WriteHeader(200) // final chunk returns the driveItem
			fmt.Fprint(w, `{"id":"item-1"}`)
		default:
			t.Errorf("unexpected call %s %s", r.Method, r.URL.Path)
		}
	}))
	srvURL = srv.URL

	// 5 MB forces the chunked path (3.2 MB chunks -> 2 chunks).
	local := filepath.Join(t.TempDir(), "big.bin")
	if err := os.WriteFile(local, make([]byte, 5<<20), 0o644); err != nil {
		t.Fatal(err)
	}

	out, err := c.UploadFile(context.Background(), "/users/u/drive/root:/big.bin", local, nil)
	if err != nil {
		t.Fatalf("upload must survive a 429 chunk: %v", err)
	}
	if !strings.Contains(string(out), "item-1") {
		t.Fatalf("final driveItem not returned: %s", out)
	}
	if throttled.Load() != 1 || chunkPuts.Load() != 3 {
		t.Fatalf("want 1 throttle + 3 chunk PUTs total, got throttled=%d puts=%d", throttled.Load(), chunkPuts.Load())
	}
}

// TestStalledDownloadAborts: a stream that stops producing bytes must be
// aborted by the idle watchdog, not hang forever.
func TestStalledDownloadAborts(t *testing.T) {
	old := streamIdleTimeout
	streamIdleTimeout = 150 * time.Millisecond
	t.Cleanup(func() { streamIdleTimeout = old })

	block := make(chan struct{})
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Length", "1000")
		w.WriteHeader(200)
		w.Write([]byte("partial-data"))
		w.(http.Flusher).Flush()
		<-block // stall: never send the rest
	}))
	// Unblock the handler BEFORE the server's Close cleanup runs (cleanups are
	// LIFO; Close waits for in-flight handlers and would deadlock otherwise).
	t.Cleanup(func() { close(block) })
	c.maxRetries = 0 // one attempt is enough to prove the abort

	start := time.Now()
	err := c.DownloadItem(context.Background(), "/users/u/drive/items/i1", filepath.Join(t.TempDir(), "f.bin"), nil)
	if err == nil {
		t.Fatal("stalled download must fail")
	}
	if !strings.Contains(err.Error(), "stalled") {
		t.Fatalf("want a stall error, got: %v", err)
	}
	if time.Since(start) > 3*time.Second {
		t.Fatalf("watchdog too slow: %v", time.Since(start))
	}
}

// TestSlowButProgressingDownloadCompletes: continuous slow progress must never
// trip the watchdog — only true stalls do.
func TestSlowButProgressingDownloadCompletes(t *testing.T) {
	old := streamIdleTimeout
	streamIdleTimeout = 200 * time.Millisecond
	t.Cleanup(func() { streamIdleTimeout = old })

	const parts = 12
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Length", fmt.Sprint(parts))
		w.WriteHeader(200)
		for i := 0; i < parts; i++ {
			w.Write([]byte("x")) // one byte per tick, each within the idle window
			w.(http.Flusher).Flush()
			time.Sleep(60 * time.Millisecond) // total 720ms > idle timeout
		}
	}))

	local := filepath.Join(t.TempDir(), "slow.bin")
	if err := c.DownloadItem(context.Background(), "/users/u/drive/items/i1", local, nil); err != nil {
		t.Fatalf("slow-but-progressing download must complete: %v", err)
	}
	data, _ := os.ReadFile(local)
	if len(data) != parts {
		t.Fatalf("want %d bytes, got %d", parts, len(data))
	}
}

// TestRetryAfterIsCapped: oversized Retry-After values are honored only up to
// the 60-second ceiling; smaller values pass through unchanged.
func TestRetryAfterIsCapped(t *testing.T) {
	for header, want := range map[string]time.Duration{"120": 60 * time.Second, "2": 2 * time.Second} {
		var calls atomic.Int32
		h := header
		c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
			if calls.Add(1) == 1 {
				w.Header().Set("Retry-After", h)
				w.WriteHeader(429)
				fmt.Fprint(w, `{"error":{"code":"activityLimitReached","message":"throttled"}}`)
				return
			}
			fmt.Fprint(w, `{}`)
		}))
		var slept []time.Duration
		c.sleep = func(ctx context.Context, d time.Duration) error {
			slept = append(slept, d)
			return nil
		}
		if err := c.Get(context.Background(), "/users", nil, nil); err != nil {
			t.Fatalf("Retry-After %s: %v", header, err)
		}
		if len(slept) != 1 || slept[0] != want {
			t.Fatalf("Retry-After %s: slept %v, want [%v]", header, slept, want)
		}
	}
}

// TestDownloadRetriesOn429: a throttled download retries and completes.
func TestDownloadRetriesOn429(t *testing.T) {
	var calls atomic.Int32
	c, _ := newTestClient(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if calls.Add(1) == 1 {
			w.Header().Set("Retry-After", "1")
			w.WriteHeader(429)
			fmt.Fprint(w, `{"error":{"code":"activityLimitReached","message":"throttled"}}`)
			return
		}
		w.WriteHeader(200)
		w.Write([]byte("content"))
	}))

	local := filepath.Join(t.TempDir(), "f.txt")
	if err := c.DownloadItem(context.Background(), "/users/u/drive/items/i1", local, nil); err != nil {
		t.Fatalf("download must survive a 429: %v", err)
	}
	if data, _ := os.ReadFile(local); string(data) != "content" {
		t.Fatalf("bad content: %q", data)
	}
	if calls.Load() != 2 {
		t.Fatalf("want exactly one retry, got %d calls", calls.Load())
	}
}
