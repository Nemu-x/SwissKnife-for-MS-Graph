package services

import (
	"bufio"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path"
	"path/filepath"
	goruntime "runtime"
	"strings"
	"sync"
	"time"

	"github.com/jedisct1/go-minisign"
	wrt "github.com/wailsapp/wails/v2/pkg/runtime"
)

// UpdateService checks GitHub Releases for a newer version of the app.
type UpdateService struct {
	ctx     context.Context
	version string

	mu sync.Mutex
	// release is the installer set seen by the last Check: Download only
	// fetches that asset and verifies it against that release's checksums.
	release *releaseAssets
	// verified maps a downloaded installer path to its verified SHA-256;
	// Apply re-hashes the file against it right before launching.
	verified map[string]string
}

// releaseAssets are the download URLs of one release that the update flow uses.
type releaseAssets struct {
	installerURL, installerName string
	sumsURL, sigURL             string
}

// releasesAPI is a variable so tests can point Check at a fake server.
var releasesAPI = "https://api.github.com/repos/Nemu-x/SwissKnife-for-MS-Graph/releases/latest"

const releasesPage = "https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/latest"

// Release asset names produced by the release workflow (sha256sum + minisign).
const (
	checksumsAsset = "SHA256SUMS.txt"
	signatureAsset = "SHA256SUMS.txt.minisig"
)

// releasePublicKey verifies SHA256SUMS.txt.minisig of every release. It is the
// RELEASE key (repo root minisign.pub, key id A4AD826C6D6BE486), not the pack
// signing key (packs.BuiltinKey). A variable only so tests can inject a key.
var releasePublicKey = "RWSG5GttbIKtpCqGS1IxYmyO07AbVOztAbwaWSi22mlNsVzJTXHRLH+9"

// updateHTTPClient fetches release assets; it refuses redirects that leave
// the GitHub release hosts. A variable so tests can use a TLS test client.
var updateHTTPClient = &http.Client{CheckRedirect: checkAssetRedirect}

func checkAssetRedirect(req *http.Request, via []*http.Request) error {
	if len(via) >= 10 {
		return errors.New("too many redirects")
	}
	if !allowedAssetURL(req.URL) {
		return errors.New("installer downloads are restricted to GitHub release hosts")
	}
	return nil
}

func NewUpdateService(version string) *UpdateService {
	return &UpdateService{version: version, verified: map[string]string{}}
}

func (u *UpdateService) SetAppContext(ctx context.Context) {
	u.ctx = ctx
	go cleanupStaleInstallers(os.TempDir(), 24*time.Hour)
}

// UpdateInfo is the result of an update check.
type UpdateInfo struct {
	CurrentVersion  string `json:"currentVersion"`
	LatestVersion   string `json:"latestVersion"`
	UpdateAvailable bool   `json:"updateAvailable"`
	Notes           string `json:"notes"`
	URL             string `json:"url"`
	AssetURL        string `json:"assetUrl,omitempty"` // installer download for this platform
	AssetName       string `json:"assetName,omitempty"`
	AssetSize       int64  `json:"assetSize,omitempty"`
}

// installerAssetName is this build's installer in a release: the NSIS package
// for the running architecture (amd64 or arm64).
func installerAssetName() string {
	return "-windows-" + goruntime.GOARCH + "-installer.exe"
}

// Check queries the latest GitHub release and compares versions.
func (u *UpdateService) Check() (*UpdateInfo, error) {
	info := &UpdateInfo{CurrentVersion: u.version, URL: releasesPage}

	req, err := http.NewRequestWithContext(context.Background(), http.MethodGet, releasesAPI, nil)
	if err != nil {
		return info, err
	}
	req.Header.Set("Accept", "application/vnd.github+json")
	client := &http.Client{Timeout: 15 * time.Second, Transport: updateHTTPClient.Transport,
		CheckRedirect: func(req *http.Request, via []*http.Request) error {
			// The API may redirect within itself (a renamed repository);
			// anything else goes through the release-host rules.
			if req.URL.Scheme == "https" && req.URL.Host == apiURLHost() && len(via) < 10 {
				return nil
			}
			return checkAssetRedirect(req, via)
		}}
	resp, err := client.Do(req)
	if err != nil {
		return info, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return info, nil // no releases yet or rate-limited — treat as up to date
	}
	var rel struct {
		TagName string `json:"tag_name"`
		Body    string `json:"body"`
		HTMLURL string `json:"html_url"`
		Assets  []struct {
			Name string `json:"name"`
			URL  string `json:"browser_download_url"`
			Size int64  `json:"size"`
		} `json:"assets"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&rel); err != nil {
		return info, err
	}
	info.LatestVersion = rel.TagName
	info.Notes = rel.Body
	if rel.HTMLURL != "" {
		info.URL = rel.HTMLURL
	}
	info.UpdateAvailable = isNewer(rel.TagName, u.version)
	// Pick this platform's installer asset for the in-app update flow, plus
	// the signed checksums of the same release it is verified against.
	ra := &releaseAssets{}
	suffix := installerAssetName()
	var assetSize int64
	for _, a := range rel.Assets {
		switch {
		case strings.HasSuffix(a.Name, suffix) && ra.installerURL == "":
			ra.installerURL, ra.installerName, assetSize = a.URL, a.Name, a.Size
		case a.Name == checksumsAsset:
			ra.sumsURL = a.URL
		case a.Name == signatureAsset:
			ra.sigURL = a.URL
		}
	}
	u.mu.Lock()
	u.release = ra
	u.mu.Unlock()
	// The in-app installer flow only exists for the Windows NSIS package.
	if goruntime.GOOS == "windows" {
		info.AssetURL, info.AssetName, info.AssetSize = ra.installerURL, ra.installerName, assetSize
	}
	return info, nil
}

// allowedAssetHosts are the only origins Download will fetch from: GitHub
// release pages and the CDN GitHub redirects release assets to.
var allowedAssetHosts = map[string]bool{
	"github.com":                           true,
	"objects.githubusercontent.com":        true,
	"release-assets.githubusercontent.com": true,
}

func allowedAssetURL(u *url.URL) bool {
	return u != nil && u.Scheme == "https" && u.User == nil && allowedAssetHosts[u.Hostname()]
}

// downloadIdleTimeout aborts the installer download when no bytes arrive for
// this long (a stall, not slowness — each received chunk resets it).
const downloadIdleTimeout = 60 * time.Second

// Download streams the release installer into the OS temp directory, emitting
// "update:progress" events, and returns the local path. Only the installer of
// the last checked release is accepted, only from GitHub release hosts; the
// byte count is checked against the asset size and the file is verified
// against the release's minisign-signed SHA256SUMS.txt before it can be run.
func (u *UpdateService) Download(assetURL, name string, size int64) (string, error) {
	if assetURL == "" || name == "" {
		return "", errors.New("no installer asset available for this platform")
	}
	parsed, err := url.Parse(assetURL)
	if err != nil || !allowedAssetURL(parsed) {
		return "", errors.New("installer downloads are restricted to GitHub release hosts")
	}
	u.mu.Lock()
	rel := u.release
	u.mu.Unlock()
	if rel == nil || rel.installerURL != assetURL || rel.installerName != name {
		return "", errors.New("this installer is not part of the last update check — check for updates again")
	}

	// Cancellable via app shutdown; an idle watchdog bounds true stalls while
	// slow-but-progressing downloads keep going (each chunk resets the timer).
	parent := u.ctx
	if parent == nil {
		parent = context.Background()
	}
	ctx, cancel := context.WithCancel(parent)
	defer cancel()

	// The signed checksums come first: a release without a valid signature is
	// refused before the installer is even downloaded.
	sum, err := expectedInstallerHash(ctx, rel)
	if err != nil {
		return "", err
	}

	watchdog := time.AfterFunc(downloadIdleTimeout, cancel)
	defer watchdog.Stop()

	req, err := http.NewRequestWithContext(ctx, http.MethodGet, assetURL, nil)
	if err != nil {
		return "", err
	}
	resp, err := updateHTTPClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("download failed: HTTP %d", resp.StatusCode)
	}

	local := filepath.Join(os.TempDir(), filepath.Base(name))
	f, err := os.Create(local)
	if err != nil {
		return "", err
	}
	defer f.Close()

	total := resp.ContentLength
	if total <= 0 {
		total = size
	}
	h := sha256.New()
	var done int64
	buf := make([]byte, 256*1024)
	for {
		n, rerr := resp.Body.Read(buf)
		if n > 0 {
			watchdog.Reset(downloadIdleTimeout)
			if _, werr := f.Write(buf[:n]); werr != nil {
				_ = f.Close()
				_ = os.Remove(local)
				return "", werr
			}
			h.Write(buf[:n])
			done += int64(n)
			emitEvent(u.ctx, "update:progress", map[string]any{"done": done, "total": total})
		}
		if rerr == io.EOF {
			break
		}
		if rerr != nil {
			_ = f.Close()
			_ = os.Remove(local)
			return "", rerr
		}
	}
	if size > 0 && done != size {
		_ = f.Close()
		_ = os.Remove(local)
		return "", fmt.Errorf("incomplete download: got %d of %d bytes", done, size)
	}
	if got := hex.EncodeToString(h.Sum(nil)); got != sum {
		_ = f.Close()
		_ = os.Remove(local)
		return "", errors.New("the downloaded installer does not match the signed release checksum — it was not run")
	}
	if err := f.Close(); err != nil {
		_ = os.Remove(local)
		return "", err
	}
	u.mu.Lock()
	u.verified[filepath.Clean(local)] = sum
	u.mu.Unlock()
	return local, nil
}

// expectedInstallerHash fetches SHA256SUMS.txt and its minisign signature from
// the release, verifies the signature with the release key and returns the
// installer's expected SHA-256 (lowercase hex).
func expectedInstallerHash(ctx context.Context, rel *releaseAssets) (string, error) {
	if rel.sumsURL == "" || rel.sigURL == "" {
		return "", errors.New("the release has no signed checksums (SHA256SUMS.txt.minisig) — download the installer from the releases page and verify it manually")
	}
	sums, err := fetchSmall(ctx, rel.sumsURL, 1<<20)
	if err != nil {
		return "", fmt.Errorf("fetch %s: %w", checksumsAsset, err)
	}
	sigText, err := fetchSmall(ctx, rel.sigURL, 16<<10)
	if err != nil {
		return "", fmt.Errorf("fetch %s: %w", signatureAsset, err)
	}
	if err := verifyChecksums(sums, sigText); err != nil {
		return "", err
	}
	sum := findChecksum(sums, rel.installerName)
	if sum == "" {
		return "", fmt.Errorf("%s has no entry for %s", checksumsAsset, rel.installerName)
	}
	return sum, nil
}

// verifyChecksums checks the minisign signature of the checksum file with the
// embedded release key (legacy and prehashed signatures are both accepted).
func verifyChecksums(sums, sigText []byte) error {
	pk, err := minisign.NewPublicKey(releasePublicKey)
	if err != nil {
		return fmt.Errorf("release key: %w", err)
	}
	sig, err := minisign.DecodeSignature(string(sigText))
	if err != nil {
		return fmt.Errorf("the release checksum signature is malformed: %w", err)
	}
	if ok, err := pk.Verify(sums, sig); err != nil || !ok {
		return errors.New("the release checksum signature is invalid — the update was refused")
	}
	return nil
}

// findChecksum returns the hash of name from sha256sum output ("<hex>  name"
// or "<hex> *name" in binary mode), or "".
func findChecksum(sums []byte, name string) string {
	sc := bufio.NewScanner(bytes.NewReader(sums))
	for sc.Scan() {
		fields := strings.Fields(sc.Text())
		if len(fields) != 2 || strings.TrimPrefix(fields[1], "*") != name {
			continue
		}
		sum := strings.ToLower(fields[0])
		if b, err := hex.DecodeString(sum); err == nil && len(b) == sha256.Size {
			return sum
		}
	}
	return ""
}

// fetchSmall GETs a small release asset (bounded by limit bytes).
func fetchSmall(ctx context.Context, assetURL string, limit int64) ([]byte, error) {
	parsed, err := url.Parse(assetURL)
	if err != nil || !allowedAssetURL(parsed) {
		return nil, errors.New("release assets are restricted to GitHub release hosts")
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, assetURL, nil)
	if err != nil {
		return nil, err
	}
	resp, err := updateHTTPClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("HTTP %d", resp.StatusCode)
	}
	b, err := io.ReadAll(io.LimitReader(resp.Body, limit+1))
	if err != nil {
		return nil, err
	}
	if int64(len(b)) > limit {
		return nil, errors.New("file is unexpectedly large")
	}
	return b, nil
}

// ErrUpdateDeclined is returned by Apply when the user dismisses the UAC
// prompt. It crosses the Wails boundary as an OpError envelope with code
// "update_declined" so the frontend can tell it apart from real failures.
var ErrUpdateDeclined = &OpError{Code: "update_declined", Message: "update cancelled: administrator approval was declined"}

// checkInstallerPath applies the guards shared by Apply and RevealInstaller:
// only a release installer executable inside the OS temp directory (i.e. what
// Download produced) is accepted. Windows-only. Returns the cleaned path.
func checkInstallerPath(installerPath string) (string, error) {
	if goruntime.GOOS != "windows" {
		return "", errors.New("in-app update is available on Windows only — use the releases page")
	}
	clean := filepath.Clean(installerPath)
	if !strings.HasSuffix(strings.ToLower(clean), "-installer.exe") {
		return "", errors.New("not a release installer executable")
	}
	if filepath.Dir(clean) != filepath.Clean(os.TempDir()) {
		return "", errors.New("installer must come from the update download location")
	}
	if _, err := os.Stat(clean); err != nil {
		return "", err
	}
	return clean, nil
}

// Apply launches the downloaded installer silently (NSIS /S) through the shell
// with the "runas" verb — the per-machine installer needs elevation, so one UAC
// prompt appears — and quits the app so the installer can replace the
// executable. A declined prompt returns ErrUpdateDeclined and leaves the app
// running; the app only quits after a successful launch.
//
// Only an installer Download verified is run, and it is re-hashed right before
// the launch through a handle that denies writers (see openInstallerLocked),
// so the file cannot be swapped between the check and the launch.
func (u *UpdateService) Apply(installerPath string) error {
	clean, err := checkInstallerPath(installerPath)
	if err != nil {
		return err
	}
	u.mu.Lock()
	want := u.verified[clean]
	u.mu.Unlock()
	if want == "" {
		return errors.New("this installer was not verified against the signed release checksums — download the update again")
	}
	f, err := openInstallerLocked(clean)
	if err != nil {
		return err
	}
	defer f.Close()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return err
	}
	if hex.EncodeToString(h.Sum(nil)) != want {
		return errors.New("the installer changed after it was verified — it was not run; download the update again")
	}
	if err := launchElevated(clean); err != nil {
		return err
	}
	// Give the frontend a beat to render the "updating" state, then exit —
	// the installer stops any stragglers and replaces the files.
	go func() {
		time.Sleep(700 * time.Millisecond)
		if u.ctx != nil {
			wrt.Quit(u.ctx)
		}
	}()
	return nil
}

// RevealInstaller opens the folder containing the downloaded installer with the
// file selected, so the user can run it by hand when the in-app launch failed.
// Same path guards as Apply.
func (u *UpdateService) RevealInstaller(installerPath string) error {
	clean, err := checkInstallerPath(installerPath)
	if err != nil {
		return err
	}
	return revealInFolder(clean)
}

// OpenReleasesPage opens a page of the project on GitHub in the system
// browser. The URL comes from the webview, so anything other than an https
// github.com/Nemu-x/ link is replaced by the default releases page.
func (u *UpdateService) OpenReleasesPage(pageURL string) {
	pageURL = safeReleasesURL(pageURL)
	if u.ctx != nil {
		wrt.BrowserOpenURL(u.ctx, pageURL)
	}
}

// apiURLHost is the host of releasesAPI.
func apiURLHost() string {
	if u, err := url.Parse(releasesAPI); err == nil {
		return u.Host
	}
	return ""
}

func safeReleasesURL(pageURL string) string {
	p, err := url.Parse(strings.TrimSpace(pageURL))
	if err != nil || p.Scheme != "https" || p.User != nil || !strings.EqualFold(p.Host, "github.com") ||
		!strings.HasPrefix(path.Clean(p.Path), "/Nemu-x/") {
		return releasesPage
	}
	return p.String()
}

// cleanupStaleInstallers removes old updater downloads from dir: installers
// older than maxAge. Files still open (a running installer) fail to delete on
// Windows and are skipped; all errors are ignored.
func cleanupStaleInstallers(dir string, maxAge time.Duration) {
	matches, err := filepath.Glob(filepath.Join(dir, "SwissKnifeGraph-*-installer.exe"))
	if err != nil {
		return
	}
	cutoff := time.Now().Add(-maxAge)
	for _, m := range matches {
		fi, err := os.Lstat(m)
		if err != nil || !fi.Mode().IsRegular() || fi.ModTime().After(cutoff) {
			continue
		}
		_ = os.Remove(m)
	}
}

// isNewer compares dotted version tags (e.g. v1.2.0). Non-numeric or dev builds
// are treated as "not newer" to avoid false prompts.
func isNewer(latest, current string) bool {
	l := parseVersion(latest)
	c := parseVersion(current)
	if l == nil || c == nil {
		return false
	}
	for i := 0; i < 3; i++ {
		if l[i] != c[i] {
			return l[i] > c[i]
		}
	}
	return false
}

func parseVersion(v string) []int {
	v = strings.TrimPrefix(strings.TrimSpace(v), "v")
	// strip pre-release suffix (e.g. 0.2.0-dev)
	if i := strings.IndexAny(v, "-+"); i >= 0 {
		v = v[:i]
	}
	parts := strings.Split(v, ".")
	if len(parts) == 0 {
		return nil
	}
	out := make([]int, 3)
	for i := 0; i < 3 && i < len(parts); i++ {
		n := 0
		for _, r := range parts[i] {
			if r < '0' || r > '9' {
				return nil
			}
			n = n*10 + int(r-'0')
		}
		out[i] = n
	}
	return out
}
