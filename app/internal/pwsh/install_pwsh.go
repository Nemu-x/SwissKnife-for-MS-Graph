package pwsh

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"sync/atomic"
	"time"
	"unicode/utf16"
)

// Installing PowerShell 7 from the app: the version comes from Microsoft's
// release metadata, the package from the official release (where Microsoft
// publishes it), checked against the published SHA-256 before anything runs;
// the operating system's own installer then asks for admin rights once.

// InstallProgress reports a stage (download, verify, install) and a percent
// (-1 when unknown).
type InstallProgress func(stage string, pct int)

// ErrDeclined: the operator said no to the admin prompt.
var ErrDeclined = errors.New("the installation was cancelled at the administrator prompt")

// DocsURL is Microsoft's page for installing PowerShell by hand.
const DocsURL = "https://learn.microsoft.com/powershell/scripting/install/installing-powershell"

// InstallMethod says how this system installs PowerShell.
type InstallMethod struct {
	Auto     bool     `json:"auto"`               // the app can install it
	How      string   `json:"how"`                // msi | pkg | snap | manual
	Commands []string `json:"commands,omitempty"` // to run by hand (manual)
	URL      string   `json:"url"`                // manual download page
}

// Overridable in tests.
var (
	releaseInfoURL = "https://aka.ms/pwsh-buildinfo-stable"
	downloadBase   = "https://github.com/PowerShell/PowerShell/releases/download/"
	// No overall timeout: a slow package download is fine while it moves;
	// the small metadata requests get their own deadline.
	httpClient = &http.Client{}
	// packageIdleTimeout aborts the package download when no bytes arrive
	// for this long (a stall, not slowness — each chunk resets it), so a dead
	// connection ends the install instead of hanging it.
	packageIdleTimeout = 60 * time.Second
)

const maxPackage = 400 << 20

func latestVersion(ctx context.Context) (string, error) {
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	req, _ := http.NewRequestWithContext(ctx, http.MethodGet, releaseInfoURL, nil)
	resp, err := httpClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("release information: %w", err)
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("release information: %s", resp.Status)
	}
	var info struct {
		ReleaseTag string `json:"ReleaseTag"`
	}
	if err := json.NewDecoder(io.LimitReader(resp.Body, 1<<16)).Decode(&info); err != nil || !stableTag(info.ReleaseTag) {
		return "", fmt.Errorf("release information: unexpected answer %q (%v)", info.ReleaseTag, err)
	}
	return strings.TrimPrefix(info.ReleaseTag, "v"), nil
}

// stableTagRe is a stable release tag (no preview/rc suffix). The version
// also becomes part of the download URL, so nothing else is accepted.
var stableTagRe = regexp.MustCompile(`^v([0-9]{1,3})\.[0-9]{1,3}\.[0-9]{1,3}$`)

// stableTag accepts PowerShell 7 or any later major version.
func stableTag(tag string) bool {
	m := stableTagRe.FindStringSubmatch(tag)
	if m == nil {
		return false
	}
	major, err := strconv.Atoi(m[1])
	return err == nil && major >= 7
}

// decodeText reads the hashes file, which is published as UTF-16.
func decodeText(b []byte) string {
	if len(b) >= 2 && b[0] == 0xFF && b[1] == 0xFE {
		u := make([]uint16, (len(b)-2)/2)
		for i := range u {
			u[i] = binary.LittleEndian.Uint16(b[2+2*i:])
		}
		return string(utf16.Decode(u))
	}
	return string(bytes.TrimPrefix(b, []byte{0xEF, 0xBB, 0xBF}))
}

// publishedHash finds the SHA-256 the release lists for file.
func publishedHash(ctx context.Context, version, file string) (string, error) {
	ctx, cancel := context.WithTimeout(ctx, time.Minute)
	defer cancel()
	req, _ := http.NewRequestWithContext(ctx, http.MethodGet, downloadBase+"v"+version+"/hashes.sha256", nil)
	resp, err := httpClient.Do(req)
	if err != nil {
		return "", fmt.Errorf("checksums: %w", err)
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("checksums: %s", resp.Status)
	}
	b, err := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
	if err != nil {
		return "", err
	}
	for _, line := range strings.Split(decodeText(b), "\n") {
		f := strings.Fields(strings.TrimSpace(line))
		if len(f) == 2 && strings.EqualFold(strings.TrimPrefix(f[1], "*"), file) && len(f[0]) == 64 {
			return strings.ToLower(f[0]), nil
		}
	}
	return "", fmt.Errorf("checksums: %s is not listed", file)
}

// fetch downloads file of the latest release into a new temporary folder and
// checks it; name builds the file name from the version. It returns the
// path and the published SHA-256 (to check again right before use).
func fetch(ctx context.Context, name func(version string) string, progress InstallProgress) (path, sum string, err error) {
	if progress == nil {
		progress = func(string, int) {}
	}
	progress("download", -1)
	version, err := latestVersion(ctx)
	if err != nil {
		return "", "", err
	}
	file := name(version)
	want, err := publishedHash(ctx, version, file)
	if err != nil {
		return "", "", err
	}
	dir, err := os.MkdirTemp("", "skg-pwsh-")
	if err != nil {
		return "", "", err
	}
	path = filepath.Join(dir, file)
	defer func() {
		if err != nil {
			_ = os.RemoveAll(dir)
		}
	}()
	dctx, cancel := context.WithCancel(ctx)
	defer cancel()
	var stalled atomic.Bool
	watchdog := time.AfterFunc(packageIdleTimeout, func() { stalled.Store(true); cancel() })
	defer watchdog.Stop()
	defer func() {
		if err != nil && stalled.Load() {
			err = fmt.Errorf("download: no data received for %s — the connection stalled", packageIdleTimeout)
		}
	}()
	req, _ := http.NewRequestWithContext(dctx, http.MethodGet, downloadBase+"v"+version+"/"+file, nil)
	resp, err := httpClient.Do(req)
	if err != nil {
		return "", "", fmt.Errorf("download: %w", err)
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		return "", "", fmt.Errorf("download: %s", resp.Status)
	}
	out, err := os.OpenFile(path, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0o600)
	if err != nil {
		return "", "", err
	}
	h := sha256.New()
	pr := &progressReader{r: io.LimitReader(resp.Body, maxPackage+1), total: resp.ContentLength, report: func(p int) { progress("download", p) },
		onData: func() { watchdog.Reset(packageIdleTimeout) }}
	n, err := io.Copy(io.MultiWriter(out, h), pr)
	if err == nil && n > maxPackage {
		err = fmt.Errorf("the package is larger than %d MB — not installing it", maxPackage>>20)
	}
	if cerr := out.Close(); err == nil {
		err = cerr
	}
	if err != nil {
		return "", "", fmt.Errorf("download: %w", err)
	}
	progress("verify", -1)
	if got := hex.EncodeToString(h.Sum(nil)); got != want {
		return "", "", fmt.Errorf("the downloaded %s does not match its published checksum — not installing it", file)
	}
	return path, want, nil
}

type progressReader struct {
	r      io.Reader
	n      int64
	total  int64
	last   int
	report func(int)
	onData func() // called whenever bytes arrive (resets the idle watchdog)
}

func (p *progressReader) Read(b []byte) (int, error) {
	n, err := p.r.Read(b)
	if n > 0 && p.onData != nil {
		p.onData()
	}
	p.n += int64(n)
	if p.total > 0 {
		if pct := int(p.n * 100 / p.total); pct != p.last {
			p.last = pct
			p.report(pct)
		}
	}
	return n, err
}

// InstallPowerShell installs PowerShell 7 the way this system does it.
func InstallPowerShell(ctx context.Context, progress InstallProgress) error {
	if progress == nil {
		progress = func(string, int) {}
	}
	return installPlatform(ctx, progress)
}
