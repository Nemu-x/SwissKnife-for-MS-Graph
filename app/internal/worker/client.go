package worker

import (
	"bytes"
	"context"
	"crypto/tls"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"

	"swissknife-app/internal/pwsh"
)

// Remote is a paired worker as the client sees it.
type Remote struct {
	Addr     string // host:port
	WorkerFP string
	http     *http.Client
}

// NewRemote builds a client that talks only to the pinned worker
// certificate and presents its own.
func NewRemote(addr, workerFP string, id *Identity) *Remote {
	cfg := &tls.Config{
		Certificates: []tls.Certificate{id.Cert},
		MinVersion:   tls.VersionTLS13,
		// The pinned fingerprint is the trust: no CA, no host name.
		InsecureSkipVerify:    true, //nolint:gosec // verified by VerifyPeerCertificate
		VerifyPeerCertificate: pinned(func(fp string) bool { return fp == workerFP }),
	}
	return &Remote{Addr: addr, WorkerFP: workerFP, http: &http.Client{
		Transport: &http.Transport{TLSClientConfig: cfg, ForceAttemptHTTP2: true},
		Timeout:   10 * time.Minute, // a mailbox search can take a while
	}}
}

func (r *Remote) url(path string) string { return "https://" + r.Addr + path }

// Health asks the worker what it runs.
func (r *Remote) Health(ctx context.Context) (*Health, error) {
	req, _ := http.NewRequestWithContext(ctx, http.MethodGet, r.url("/health"), nil)
	resp, err := r.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("worker: %s", resp.Status)
	}
	var h Health
	return &h, json.NewDecoder(io.LimitReader(resp.Body, 1<<16)).Decode(&h)
}

// Invoke runs one cmdlet on the worker; a PowerShell failure comes back as
// *pwsh.Error, like a local call.
func (r *Remote) Invoke(ctx context.Context, in InvokeRequest) ([]json.RawMessage, error) {
	body, _ := json.Marshal(in)
	req, _ := http.NewRequestWithContext(ctx, http.MethodPost, r.url("/invoke"), bytes.NewReader(body))
	req.Header.Set("Content-Type", "application/json")
	resp, err := r.http.Do(req)
	if err != nil {
		return nil, fmt.Errorf("worker %s: %w", r.Addr, err)
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		b, _ := io.ReadAll(io.LimitReader(resp.Body, 512))
		return nil, fmt.Errorf("worker: %s %s", resp.Status, strings.TrimSpace(string(b)))
	}
	var out InvokeResponse
	if err := json.NewDecoder(io.LimitReader(resp.Body, 64<<20)).Decode(&out); err != nil {
		return nil, err
	}
	if out.Error != nil {
		return nil, out.Error
	}
	return out.Data, nil
}

// Pair connects with a one-time code and returns the worker's fingerprint
// to pin, plus its name.
func Pair(ctx context.Context, addr, code, name string, id *Identity) (fp, workerName string, err error) {
	var seen string
	cfg := &tls.Config{
		Certificates: []tls.Certificate{id.Cert},
		MinVersion:   tls.VersionTLS13,
		// Not trusted yet: the code proves both sides below.
		InsecureSkipVerify: true, //nolint:gosec // bound by the pairing proofs
		VerifyConnection: func(cs tls.ConnectionState) error {
			if len(cs.PeerCertificates) == 0 {
				return errors.New("no worker certificate")
			}
			seen = Fingerprint(cs.PeerCertificates[0].Raw)
			return nil
		},
	}
	hc := &http.Client{Transport: &http.Transport{TLSClientConfig: cfg, DisableKeepAlives: true}, Timeout: 30 * time.Second}
	// The proof needs the fingerprint seen in the handshake: handshake first.
	conn, err := tls.Dial("tcp", addr, cfg)
	if err != nil {
		return "", "", fmt.Errorf("worker %s: %w", addr, err)
	}
	_ = conn.Close()
	first := seen
	body, _ := json.Marshal(map[string]string{"name": name, "proof": ClientProof(code, id.Fingerprint, first)})
	req, _ := http.NewRequestWithContext(ctx, http.MethodPost, "https://"+addr+"/pair", bytes.NewReader(body))
	resp, err := hc.Do(req)
	if err != nil {
		return "", "", err
	}
	defer func() { _ = resp.Body.Close() }()
	if seen != first {
		return "", "", errors.New("the worker's certificate changed during pairing")
	}
	if resp.StatusCode != http.StatusOK {
		b, _ := io.ReadAll(io.LimitReader(resp.Body, 512))
		return "", "", fmt.Errorf("pairing refused: %s", strings.TrimSpace(string(b)))
	}
	var out struct{ Proof, Name string }
	if err := json.NewDecoder(io.LimitReader(resp.Body, 4096)).Decode(&out); err != nil {
		return "", "", err
	}
	if out.Proof != WorkerProof(code, id.Fingerprint, first) {
		return "", "", errors.New("the worker could not prove the code — do not trust this connection")
	}
	return first, out.Name, nil
}

// ensure the pwsh error type crosses the wire intact
var _ error = (*pwsh.Error)(nil)
