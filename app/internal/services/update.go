package services

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	goruntime "runtime"
	"strings"
	"time"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"
)

// UpdateService checks GitHub Releases for a newer version of the app.
type UpdateService struct {
	ctx     context.Context
	version string
}

const releasesAPI = "https://api.github.com/repos/Nemu-x/SwissKnife-for-MS-Graph/releases/latest"
const releasesPage = "https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/latest"

func NewUpdateService(version string) *UpdateService {
	return &UpdateService{version: version}
}

func (u *UpdateService) SetAppContext(ctx context.Context) { u.ctx = ctx }

// UpdateInfo is the result of an update check.
type UpdateInfo struct {
	CurrentVersion  string `json:"currentVersion"`
	LatestVersion   string `json:"latestVersion"`
	UpdateAvailable bool   `json:"updateAvailable"`
	Notes           string `json:"notes"`
	URL             string `json:"url"`
	AssetURL        string `json:"assetUrl,omitempty"`  // installer download for this platform
	AssetName       string `json:"assetName,omitempty"`
	AssetSize       int64  `json:"assetSize,omitempty"`
}

// Check queries the latest GitHub release and compares versions.
func (u *UpdateService) Check() (*UpdateInfo, error) {
	info := &UpdateInfo{CurrentVersion: u.version, URL: releasesPage}

	req, err := http.NewRequestWithContext(context.Background(), http.MethodGet, releasesAPI, nil)
	if err != nil {
		return info, err
	}
	req.Header.Set("Accept", "application/vnd.github+json")
	client := &http.Client{Timeout: 15 * time.Second}
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
	// Pick this platform's installer asset for the in-app update flow.
	if goruntime.GOOS == "windows" {
		for _, a := range rel.Assets {
			if strings.HasSuffix(a.Name, "-installer.exe") {
				info.AssetURL, info.AssetName, info.AssetSize = a.URL, a.Name, a.Size
				break
			}
		}
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

// downloadIdleTimeout aborts the installer download when no bytes arrive for
// this long (a stall, not slowness — each received chunk resets it).
const downloadIdleTimeout = 60 * time.Second

// Download streams the release installer into the OS temp directory, emitting
// "update:progress" events, and returns the local path. Only GitHub release
// hosts are accepted, and the byte count is verified against the release asset
// size before anything gets executed.
func (u *UpdateService) Download(assetURL, name string, size int64) (string, error) {
	if assetURL == "" || name == "" {
		return "", errors.New("no installer asset available for this platform")
	}
	parsed, err := url.Parse(assetURL)
	if err != nil || parsed.Scheme != "https" || !allowedAssetHosts[parsed.Hostname()] {
		return "", errors.New("installer downloads are restricted to GitHub release hosts")
	}

	// Cancellable via app shutdown; an idle watchdog bounds true stalls while
	// slow-but-progressing downloads keep going (each chunk resets the timer).
	parent := u.ctx
	if parent == nil {
		parent = context.Background()
	}
	ctx, cancel := context.WithCancel(parent)
	defer cancel()
	watchdog := time.AfterFunc(downloadIdleTimeout, cancel)
	defer watchdog.Stop()

	req, err := http.NewRequestWithContext(ctx, http.MethodGet, assetURL, nil)
	if err != nil {
		return "", err
	}
	resp, err := http.DefaultClient.Do(req)
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
	var done int64
	buf := make([]byte, 256*1024)
	for {
		n, rerr := resp.Body.Read(buf)
		if n > 0 {
			watchdog.Reset(downloadIdleTimeout)
			if _, werr := f.Write(buf[:n]); werr != nil {
				_ = os.Remove(local)
				return "", werr
			}
			done += int64(n)
			emitEvent(u.ctx, "update:progress", map[string]any{"done": done, "total": total})
		}
		if rerr == io.EOF {
			break
		}
		if rerr != nil {
			_ = os.Remove(local)
			return "", rerr
		}
	}
	if size > 0 && done != size {
		_ = os.Remove(local)
		return "", fmt.Errorf("incomplete download: got %d of %d bytes", done, size)
	}
	return local, nil
}

// Apply launches the downloaded installer silently and quits the app so the
// installer can replace the executable. Windows-only (NSIS /S). Only installer
// executables inside the OS temp directory (i.e. what Download produced) run.
func (u *UpdateService) Apply(installerPath string) error {
	if goruntime.GOOS != "windows" {
		return errors.New("in-app update is available on Windows only — use the releases page")
	}
	clean := filepath.Clean(installerPath)
	if !strings.HasSuffix(strings.ToLower(clean), "-installer.exe") {
		return errors.New("not a release installer executable")
	}
	if filepath.Dir(clean) != filepath.Clean(os.TempDir()) {
		return errors.New("installer must come from the update download location")
	}
	if _, err := os.Stat(clean); err != nil {
		return err
	}
	cmd := exec.Command(clean, "/S")
	if err := cmd.Start(); err != nil {
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

// OpenReleasesPage opens the releases page in the system browser.
func (u *UpdateService) OpenReleasesPage(pageURL string) {
	if pageURL == "" {
		pageURL = releasesPage
	}
	if u.ctx != nil {
		wrt.BrowserOpenURL(u.ctx, pageURL)
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
