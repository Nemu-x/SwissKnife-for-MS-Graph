package graphapi

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"time"
)

// simpleUploadLimit is the size up to which Graph allows a plain PUT .../content.
const simpleUploadLimit = 4 << 20 // 4 MiB

// streamIdleTimeout aborts a streaming download when no bytes arrive for this
// long — slow-but-progressing transfers are never killed, stalls are. It is a
// variable so tests can shorten it.
var streamIdleTimeout = 60 * time.Second

// chunkAttemptTimeout bounds a single chunk-upload attempt; a chunk this small
// (3.2 MB) that takes longer is effectively stalled and worth retrying.
const chunkAttemptTimeout = 2 * time.Minute

// errStalled marks an idle-watchdog abort, distinguished from operator cancel.
var errStalled = errors.New("transfer stalled — no data received")

// Progress is a transfer progress callback (bytes done / total).
type Progress func(done, total int64)

// DownloadItem downloads content at itemPath (e.g. /users/u/drive/items/{id})
// into localPath, streaming to disk. Throttling (429) and transport failures
// are retried like any idempotent request; the whole file restarts on retry.
func (c *Client) DownloadItem(ctx context.Context, itemPath, localPath string, progress Progress) error {
	u := c.baseURL + "/" + trimSlash(itemPath) + "/content"
	return c.withRetry(ctx, http.MethodGet, func() (time.Duration, error) {
		return c.downloadOnce(ctx, u, localPath, progress)
	})
}

func (c *Client) downloadOnce(parent context.Context, u, localPath string, progress Progress) (time.Duration, error) {
	// Idle watchdog: cancels the request context when no read happens for
	// streamIdleTimeout; each successful read pushes the deadline forward.
	ctx, cancel := context.WithCancel(parent)
	defer cancel()
	watchdog := time.AfterFunc(streamIdleTimeout, cancel)
	defer watchdog.Stop()
	// stalled distinguishes a watchdog abort from an operator/app cancel.
	stalled := func(err error) error {
		if parent.Err() == nil && ctx.Err() != nil {
			return errStalled
		}
		return err
	}

	token, err := c.tokens.Token(ctx)
	if err != nil {
		return 0, err
	}
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return 0, err
	}
	req.Header.Set("Authorization", "Bearer "+token)

	resp, err := c.http.Do(req)
	if err != nil {
		return 0, stalled(err)
	}
	defer resp.Body.Close()

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		raw, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
		return parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, raw)
	}

	if err := os.MkdirAll(filepath.Dir(localPath), 0o755); err != nil {
		return 0, err
	}
	f, err := os.Create(localPath)
	if err != nil {
		return 0, err
	}
	defer f.Close()

	total := resp.ContentLength
	var done int64
	buf := make([]byte, 256*1024)
	for {
		n, rerr := resp.Body.Read(buf)
		if n > 0 {
			watchdog.Reset(streamIdleTimeout)
			if _, werr := f.Write(buf[:n]); werr != nil {
				return 0, werr
			}
			done += int64(n)
			if progress != nil {
				progress(done, total)
			}
		}
		if rerr == io.EOF {
			return 0, nil
		}
		if rerr != nil {
			return 0, stalled(rerr)
		}
	}
}

// UploadFile uploads a local file to the drive path uploadRoot
// (e.g. /users/u/drive/root:/docs/report.pdf). Small files via a plain PUT,
// large files via a chunked upload session (plain PUT limit is ~4 MB).
func (c *Client) UploadFile(ctx context.Context, uploadRoot, localPath string, progress Progress) (json.RawMessage, error) {
	info, err := os.Stat(localPath)
	if err != nil {
		return nil, err
	}

	if info.Size() <= simpleUploadLimit {
		return c.uploadSimple(ctx, uploadRoot, localPath, info.Size(), progress)
	}
	return c.uploadChunked(ctx, uploadRoot, localPath, info.Size(), progress)
}

func (c *Client) uploadSimple(ctx context.Context, uploadRoot, localPath string, size int64, progress Progress) (json.RawMessage, error) {
	data, err := os.ReadFile(localPath)
	if err != nil {
		return nil, err
	}
	u := c.baseURL + "/" + trimSlash(uploadRoot) + ":/content"

	var out json.RawMessage
	err = c.withRetry(ctx, http.MethodPut, func() (time.Duration, error) {
		token, err := c.tokens.Token(ctx)
		if err != nil {
			return 0, err
		}
		req, err := http.NewRequestWithContext(ctx, http.MethodPut, u, bytes.NewReader(data))
		if err != nil {
			return 0, err
		}
		req.Header.Set("Authorization", "Bearer "+token)
		req.Header.Set("Content-Type", "application/octet-stream")

		resp, err := c.http.Do(req)
		if err != nil {
			return 0, err
		}
		defer resp.Body.Close()
		raw, _ := io.ReadAll(resp.Body)
		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			return parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, raw)
		}
		out = raw
		return 0, nil
	})
	if err != nil {
		return nil, err
	}
	if progress != nil {
		progress(size, size)
	}
	return out, nil
}

func (c *Client) uploadChunked(ctx context.Context, uploadRoot, localPath string, size int64, progress Progress) (json.RawMessage, error) {
	// 1. create the upload session
	var session struct {
		UploadURL string `json:"uploadUrl"`
	}
	body := map[string]any{
		"item": map[string]any{"@microsoft.graph.conflictBehavior": "replace"},
	}
	if err := c.Post(ctx, trimSlash(uploadRoot)+":/createUploadSession", body, &session); err != nil {
		return nil, err
	}
	if session.UploadURL == "" {
		return nil, fmt.Errorf("graph: empty uploadUrl")
	}

	f, err := os.Open(localPath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	// 2. push chunks (uploadUrl is already authorized — no Bearer needed)
	const chunk = int64(10 * 320 * 1024) // 3.2 MB, a multiple of 320 KiB
	buf := make([]byte, chunk)
	var offset int64
	for offset < size {
		n := chunk
		if size-offset < n {
			n = size - offset
		}
		if _, err := io.ReadFull(f, buf[:n]); err != nil {
			return nil, err
		}

		// Each chunk retries independently (429/5xx/transport) with a bounded
		// per-attempt deadline — a 3.2 MB chunk taking longer is stalled.
		var raw []byte
		var status int
		err := c.withRetry(ctx, http.MethodPut, func() (time.Duration, error) {
			attemptCtx, cancel := context.WithTimeout(ctx, chunkAttemptTimeout)
			defer cancel()
			req, err := http.NewRequestWithContext(attemptCtx, http.MethodPut, session.UploadURL, bytes.NewReader(buf[:n]))
			if err != nil {
				return 0, err
			}
			req.Header.Set("Content-Range", fmt.Sprintf("bytes %d-%d/%d", offset, offset+n-1, size))
			req.ContentLength = n

			resp, err := c.http.Do(req)
			if err != nil {
				// A per-attempt deadline expiry must stay retryable: report it
				// as a plain stall, not context.DeadlineExceeded.
				if ctx.Err() == nil && attemptCtx.Err() != nil {
					return 0, errStalled
				}
				return 0, err
			}
			body, _ := io.ReadAll(resp.Body)
			resp.Body.Close()
			if resp.StatusCode < 200 || resp.StatusCode >= 300 {
				return parseRetryAfter(resp.Header.Get("Retry-After")), parseGraphError(resp, body)
			}
			raw, status = body, resp.StatusCode
			return 0, nil
		})
		if err != nil {
			return nil, err
		}

		offset += n
		if progress != nil {
			progress(offset, size)
		}

		// the final chunk returns the driveItem
		if status == 200 || status == 201 {
			return raw, nil
		}
	}
	return nil, fmt.Errorf("graph: upload session finished without final response")
}

func trimSlash(s string) string {
	for len(s) > 0 && s[0] == '/' {
		s = s[1:]
	}
	return s
}
