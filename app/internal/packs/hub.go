package packs

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
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
	hubClient = &http.Client{Timeout: time.Minute}
	fileRe    = regexp.MustCompile(`^[A-Za-z0-9._-]+(/[A-Za-z0-9._-]+)*$`)
)

func safeRel(p string) bool {
	return fileRe.MatchString(p) && !strings.Contains(p, "..") && !strings.HasPrefix(p, ".")
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
		if nameRe.MatchString(e.Name) && safeRel(e.Path) && len(e.Digest) == 64 {
			kept = append(kept, e)
		}
	}
	idx.Packs = kept
	return &idx, nil
}

// Install downloads entry into root/<name>, replacing an older copy only
// once the new one is complete and matches the index's digest.
func Install(ctx context.Context, hub, root string, e HubEntry) (string, error) {
	if !nameRe.MatchString(e.Name) || !safeRel(e.Path) {
		return "", errors.New("invalid pack entry")
	}
	if len(e.Files) == 0 || len(e.Files) > maxFiles {
		return "", errors.New("a pack lists 1 to 50 files")
	}
	if err := os.MkdirAll(root, 0o755); err != nil {
		return "", err
	}
	tmp, err := os.MkdirTemp(root, ".hub-"+e.Name+"-")
	if err != nil {
		return "", err
	}
	defer func() { _ = os.RemoveAll(tmp) }()
	total := 0
	for _, f := range e.Files {
		if !safeRel(f) {
			return "", fmt.Errorf("invalid file name %q", f)
		}
		u := baseURL(hub) + path.Join(e.Path, f)
		if _, err := url.Parse(u); err != nil {
			return "", err
		}
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
	dir := filepath.Join(root, e.Name)
	old := dir + ".old"
	_ = os.RemoveAll(old)
	if _, err := os.Stat(dir); err == nil {
		if err := os.Rename(dir, old); err != nil {
			return "", err
		}
	}
	if err := os.Rename(tmp, dir); err != nil {
		_ = os.Rename(old, dir)
		return "", err
	}
	_ = os.RemoveAll(old)
	return dir, nil
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
		if !de.IsDir() {
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
		_, signedErr := os.Stat(filepath.Join(p.Dir, SignatureFile))
		idx.Packs = append(idx.Packs, HubEntry{Name: m.Name, Version: m.Version, Category: cat, Kind: kind, Title: title,
			Description: desc, Author: m.Author, Path: "packs/" + de.Name(), Files: names, Digest: p.Digest, Signed: signedErr == nil})
	}
	sort.Slice(idx.Packs, func(i, j int) bool { return idx.Packs[i].Name < idx.Packs[j].Name })
	return idx, nil
}
