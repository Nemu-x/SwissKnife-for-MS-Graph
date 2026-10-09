package packs

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

// The Action Hub: a public catalog of packs (index.json plus one folder per
// pack). Installing downloads the listed files and checks that they hash to
// the digest the index names; trust works as for any pack afterwards — a
// signature by a trusted key, or the operator's review.

// DefaultHub is the community catalog.
const DefaultHub = "https://raw.githubusercontent.com/Nemu-x/SwissKnife-Action-Hub/main/"

// Categories packs may declare.
var Categories = map[string]bool{"security": true, "exchange": true, "hr": true, "teams": true, "intune": true, "reports": true, "other": true}

// HubIndex is index.json.
type HubIndex struct {
	Version int        `json:"version"`
	Packs   []HubEntry `json:"packs"`
}

// HubEntry describes one pack of the catalog.
type HubEntry struct {
	Name        string            `json:"name"`
	Version     string            `json:"version"`
	Category    string            `json:"category"`
	Kind        string            `json:"kind"` // workflow | script | mixed
	Title       map[string]string `json:"title"`
	Description map[string]string `json:"description"`
	Author      string            `json:"author"`
	Path        string            `json:"path"`  // folder in the hub
	Files       []string          `json:"files"` // relative to Path
	Digest      string            `json:"digest"`
	Signed      bool              `json:"signed"`
}

const (
	maxIndex     = 1 << 20
	maxFile      = 1 << 20
	maxPackBytes = 5 << 20
	maxFiles     = 50
)

var (
	// Redirects stay on https and on the same host: the index is not signed,
	// so the transport is what keeps it the hub's.
	hubClient = &http.Client{Timeout: time.Minute, CheckRedirect: func(r *http.Request, via []*http.Request) error {
		if len(via) >= 5 || r.URL.Scheme != "https" || r.URL.Host != via[0].URL.Host {
			return errors.New("the hub redirected somewhere else — refused")
		}
		return nil
	}}
	fileRe = regexp.MustCompile(`^[A-Za-z0-9._-]+(/[A-Za-z0-9._-]+)*$`)
	// Device names Windows reserves in every folder.
	reservedRe = regexp.MustCompile(`(?i)^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$`)
)

func safeRel(p string) bool {
	if !fileRe.MatchString(p) || strings.Contains(p, "..") || strings.HasPrefix(p, ".") {
		return false
	}
	for _, seg := range strings.Split(p, "/") {
		if reservedRe.MatchString(seg) || strings.HasSuffix(seg, ".") || strings.HasPrefix(seg, ".") {
			return false
		}
	}
	return true
}

// hubMu runs one install at a time.
var hubMu sync.Mutex

// Newer reports whether version a is newer than b (dotted numbers; text
// parts compare as text). Only newer versions are offered as updates.
func Newer(a, b string) bool {
	pa, pb := strings.Split(a, "."), strings.Split(b, ".")
	for i := 0; i < len(pa) || i < len(pb); i++ {
		var x, y string
		if i < len(pa) {
			x = pa[i]
		}
		if i < len(pb) {
			y = pb[i]
		}
		nx, ex := strconv.Atoi(x)
		ny, ey := strconv.Atoi(y)
		switch {
		case ex == nil && ey == nil && nx != ny:
			return nx > ny
		case (ex != nil || ey != nil) && x != y:
			return x > y
		}
	}
	return false
}

func hubGet(ctx context.Context, u string, limit int64) ([]byte, error) {
	if !strings.HasPrefix(u, "https://") {
		return nil, errors.New("the hub must be an https address")
	}
	req, _ := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	resp, err := hubClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer func() { _ = resp.Body.Close() }()
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("%s: %s", u, resp.Status)
	}
	b, err := io.ReadAll(io.LimitReader(resp.Body, limit+1))
	if err != nil {
		return nil, err
	}
	if int64(len(b)) > limit {
		return nil, fmt.Errorf("%s is larger than allowed", u)
	}
	return b, nil
}

func baseURL(hub string) string {
	if !strings.HasSuffix(hub, "/") {
		hub += "/"
	}
	return hub
}

// FetchIndex reads the hub's index.json.
func FetchIndex(ctx context.Context, hub string) (*HubIndex, error) {
	b, err := hubGet(ctx, baseURL(hub)+"index.json", maxIndex)
	if err != nil {
		return nil, err
	}
	var idx HubIndex
	if err := json.Unmarshal(b, &idx); err != nil {
		return nil, fmt.Errorf("index.json: %w", err)
	}
	kept := idx.Packs[:0]
	for _, e := range idx.Packs {
		if nameRe.MatchString(e.Name) && !reservedRe.MatchString(e.Name) && safeRel(e.Path) && len(e.Digest) == 64 {
			if e.Kind != "workflow" && e.Kind != "script" && e.Kind != "mixed" {
				e.Kind = "mixed" // claims nothing it cannot show
			}
			if !Categories[e.Category] {
				e.Category = "other"
			}
			kept = append(kept, e)
		}
	}
	idx.Packs = kept
	return &idx, nil
}

// Install downloads entry and puts it in place of the installed pack of the
// same name (or root/<name>), only once it is complete, matches the index's
// digest and its manifest names the same pack.
func Install(ctx context.Context, hub, root string, e HubEntry) (string, error) {
	if !nameRe.MatchString(e.Name) || reservedRe.MatchString(e.Name) || !safeRel(e.Path) {
		return "", errors.New("invalid pack entry")
	}
	if len(e.Files) == 0 || len(e.Files) > maxFiles {
		return "", errors.New("a pack lists 1 to 50 files")
	}
	hubMu.Lock()
	defer hubMu.Unlock()
	if err := os.MkdirAll(root, 0o755); err != nil {
		return "", err
	}
	cleanStale(root)
	tmp, err := os.MkdirTemp(root, ".staging-"+e.Name+"-")
	if err != nil {
		return "", err
	}
	defer func() { _ = os.RemoveAll(tmp) }()
	total := 0
	seenFiles := map[string]bool{}
	for _, f := range e.Files {
		if !safeRel(f) || seenFiles[strings.ToLower(f)] {
			return "", fmt.Errorf("invalid or repeated file name %q", f)
		}
		seenFiles[strings.ToLower(f)] = true
		u := baseURL(hub) + path.Join(e.Path, f)
		b, err := hubGet(ctx, u, maxFile)
		if err != nil {
			return "", err
		}
		total += len(b)
		if total > maxPackBytes {
			return "", errors.New("the pack is larger than allowed")
		}
		dst := filepath.Join(tmp, filepath.FromSlash(f))
		if err := os.MkdirAll(filepath.Dir(dst), 0o755); err != nil {
			return "", err
		}
		if err := os.WriteFile(dst, b, 0o644); err != nil {
			return "", err
		}
	}
	d, err := Digest(tmp)
	if err != nil {
		return "", err
	}
	if d != e.Digest {
		return "", errors.New("the downloaded pack does not match the catalog — not installing it")
	}
	got := loadOne(tmp, Trust{}, trustedKeys(nil))
	if got.Manifest.Name != e.Name {
		return "", fmt.Errorf("the downloaded pack is named %q, not %q — not installing it", got.Manifest.Name, e.Name)
	}
	// Update the folder that holds this pack now; never replace a folder
	// that holds a different pack.
	dir := filepath.Join(root, e.Name)
	for _, p := range Load(root, Trust{}) {
		if p.Manifest.Name == e.Name {
			dir = p.Dir
			break
		}
	}
	if _, err := os.Stat(dir); err == nil {
		if cur := loadOne(dir, Trust{}, trustedKeys(nil)); cur.Manifest.Name != "" && cur.Manifest.Name != e.Name {
			return "", fmt.Errorf("%s holds another pack (%s) — move it first", dir, cur.Manifest.Name)
		}
	}
	trash := filepath.Join(root, ".trash-"+e.Name)
	_ = os.RemoveAll(trash)
	if _, err := os.Stat(dir); err == nil {
		if err := os.Rename(dir, trash); err != nil {
			return "", err
		}
	}
	if err := os.Rename(tmp, dir); err != nil {
		_ = os.Rename(trash, dir)
		return "", err
	}
	_ = os.RemoveAll(trash) // a held file leaves it: hidden, removed next time
	return dir, nil
}

// cleanStale removes staging and backup folders a crash or a held file left.
func cleanStale(root string) {
	entries, _ := os.ReadDir(root)
	for _, e := range entries {
		if e.IsDir() && (strings.HasPrefix(e.Name(), ".staging-") || strings.HasPrefix(e.Name(), ".trash-") || strings.HasPrefix(e.Name(), ".hub-")) {
			_ = os.RemoveAll(filepath.Join(root, e.Name()))
		}
	}
}

// BuildIndex writes the index of every pack under dir/packs (for the hub's
// maintainers: "SwissKnifeGraph pack index <hub folder>").
func BuildIndex(dir string) (*HubIndex, error) {
	root := filepath.Join(dir, "packs")
	entries, err := os.ReadDir(root)
	if err != nil {
		return nil, err
	}
	idx := &HubIndex{Version: 1, Packs: []HubEntry{}}
	for _, de := range entries {
		if !de.IsDir() || strings.HasPrefix(de.Name(), ".") {
			continue
		}
		p := loadOne(filepath.Join(root, de.Name()), Trust{}, trustedKeys(nil))
		if p.Status == Invalid && !strings.Contains(p.Error, "signature") {
			return nil, fmt.Errorf("%s: %s", de.Name(), p.Error)
		}
		_, files, err := readPack(p.Dir)
		if err != nil {
			return nil, err
		}
		var names []string
		for f := range files {
			names = append(names, f)
		}
		for _, f := range []string{DigestFile, SignatureFile} {
			if _, err := os.Stat(filepath.Join(p.Dir, f)); err == nil {
				names = append(names, f)
			}
		}
		sort.Strings(names)
		m := p.Manifest
		kind := "workflow"
		switch {
		case len(m.Actions) > 0 && len(m.Workflows) > 0:
			kind = "mixed"
		case len(m.Actions) > 0:
			kind = "script"
		}
		title := m.Title
		if len(title) == 0 {
			title = map[string]string{"en": m.Name}
		}
		desc := m.Summary
		if len(desc) == 0 && m.Description != "" {
			desc = map[string]string{"en": m.Description}
		}
		cat := m.Category
		if !Categories[cat] {
			cat = "other"
		}
		// "Signed" only when the signature verifies against the project key.
		idx.Packs = append(idx.Packs, HubEntry{Name: m.Name, Version: m.Version, Category: cat, Kind: kind, Title: title,
			Description: desc, Author: m.Author, Path: "packs/" + de.Name(), Files: names, Digest: p.Digest, Signed: p.Status == Signed})
	}
	sort.Slice(idx.Packs, func(i, j int) bool { return idx.Packs[i].Name < idx.Packs[j].Name })
	return idx, nil
}
