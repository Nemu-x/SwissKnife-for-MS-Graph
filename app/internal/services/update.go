package services

import (
	"context"
	"encoding/json"
	"net/http"
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
	return info, nil
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
