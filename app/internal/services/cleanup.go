package services

import (
	"encoding/json"
	"errors"
	"net/url"
	"sort"
	"sync/atomic"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/session"
)

// CleanupService reclaims OneDrive/SharePoint storage: duplicate files and
// version-history bloat. For SharePoint sites it scans every document library.
type CleanupService struct {
	s      *session.Session
	cancel atomic.Bool // set by CancelScan to abort an in-flight scan
}

func NewCleanupService(s *session.Session) *CleanupService { return &CleanupService{s: s} }

// CancelScan requests that the current scan stop; it returns whatever it has
// found so far rather than an error.
func (cl *CleanupService) CancelScan() { cl.cancel.Store(true) }

// emit sends a live progress line to the UI ("cleanup:progress" event).
func (cl *CleanupService) emit(stage string, done, total int) {
	wrt.EventsEmit(cl.s.Ctx(), "cleanup:progress", map[string]any{"stage": stage, "done": done, "total": total})
}

// driveBases returns the API base path(s) for the owner's drive(s).
// A user has one drive; a site can have several document libraries.
func (cl *CleanupService) driveBases(ownerType, ownerID string) ([]string, error) {
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	switch ownerType {
	case "user":
		return []string{"/users/" + url.PathEscape(ownerID) + "/drive"}, nil
	case "site":
		var resp struct {
			Value []struct {
				ID string `json:"id"`
			} `json:"value"`
		}
		if err := c.Get(cl.s.Ctx(), "/sites/"+url.PathEscape(ownerID)+"/drives", url.Values{"$select": {"id"}}, &resp); err == nil && len(resp.Value) > 0 {
			out := make([]string, 0, len(resp.Value))
			for _, d := range resp.Value {
				out = append(out, "/drives/"+url.PathEscape(d.ID))
			}
			return out, nil
		}
		return []string{"/sites/" + url.PathEscape(ownerID) + "/drive"}, nil
	}
	return nil, errors.New("ownerType must be 'user' or 'site'")
}

type fileRec struct {
	id, name, path, base string
	size                 int64
	quickXor, sha        string
}

// walkFiles collects every file across all of the owner's drives.
func (cl *CleanupService) walkFiles(ownerType, ownerID string) ([]fileRec, error) {
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	bases, err := cl.driveBases(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	var files []fileRec
	var walk func(base, itemsPath, rel string) error
	walk = func(base, itemsPath, rel string) error {
		if cl.cancel.Load() {
			return nil
		}
		items, err := c.ListAll(cl.s.Ctx(), itemsPath, nil, 0)
		if err != nil {
			return err
		}
		for _, raw := range items {
			if cl.cancel.Load() {
				return nil
			}
			var it struct {
				ID     string          `json:"id"`
				Name   string          `json:"name"`
				Size   int64           `json:"size"`
				Folder json.RawMessage `json:"folder"`
				File   struct {
					Hashes struct {
						QuickXor string `json:"quickXorHash"`
						Sha256   string `json:"sha256Hash"`
					} `json:"hashes"`
				} `json:"file"`
			}
			if json.Unmarshal(raw, &it) != nil || it.Name == "" {
				continue
			}
			relPath := it.Name
			if rel != "" {
				relPath = rel + "/" + it.Name
			}
			if it.Folder != nil {
				_ = walk(base, base+"/items/"+url.PathEscape(it.ID)+"/children", relPath)
				continue
			}
			files = append(files, fileRec{
				id: it.ID, name: it.Name, path: relPath, base: base, size: it.Size,
				quickXor: it.File.Hashes.QuickXor, sha: it.File.Hashes.Sha256,
			})
			if len(files)%100 == 0 {
				cl.emit("Scanning files", len(files), 0)
			}
		}
		return nil
	}
	for _, base := range bases {
		if err := walk(base, base+"/root/children", ""); err != nil {
			return nil, err
		}
	}
	return files, nil
}

// --- Duplicate files ---

type DupItem struct {
	Ref  string `json:"ref"` // API path used to delete this copy
	Name string `json:"name"`
	Path string `json:"path"`
}

type DupGroup struct {
	Name   string    `json:"name"`
	Size   int64     `json:"size"`
	Count  int       `json:"count"`
	Wasted int64     `json:"wasted"`
	Items  []DupItem `json:"items"`
}

// FindDuplicates groups byte-identical files (by content hash, else name+size).
func (cl *CleanupService) FindDuplicates(ownerType, ownerID string) ([]DupGroup, error) {
	cl.cancel.Store(false)
	files, err := cl.walkFiles(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	groups := map[string][]fileRec{}
	for _, f := range files {
		key := f.quickXor
		if key == "" {
			key = f.sha
		}
		if key == "" {
			key = f.name + "|" + itoa(int(f.size))
		}
		groups[key] = append(groups[key], f)
	}
	out := make([]DupGroup, 0)
	for _, recs := range groups {
		if len(recs) < 2 {
			continue
		}
		g := DupGroup{Name: recs[0].name, Size: recs[0].size, Count: len(recs), Wasted: int64(len(recs)-1) * recs[0].size}
		for _, r := range recs {
			g.Items = append(g.Items, DupItem{Ref: r.base + "/items/" + url.PathEscape(r.id), Name: r.name, Path: r.path})
		}
		out = append(out, g)
	}
	cl.emit("Done", len(files), len(files))
	sort.Slice(out, func(i, j int) bool { return out[i].Wasted > out[j].Wasted })
	return out, nil
}

// DeleteItems removes drive items by their API ref. Destructive: confirm == "DELETE".
func (cl *CleanupService) DeleteItems(refs []string, confirm string) (map[string]any, error) {
	if err := cl.s.GuardWrite(); err != nil {
		return nil, err
	}
	if confirm != "DELETE" {
		return nil, errors.New("type DELETE to confirm bulk deletion")
	}
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	deleted := 0
	failures := map[string]string{}
	for _, ref := range refs {
		if e := c.Delete(cl.s.Ctx(), ref); e != nil {
			failures[ref] = e.Error()
			continue
		}
		deleted++
	}
	cl.s.Record("cleanup.deleteItems", "", "deleted="+itoa(deleted), nil)
	return map[string]any{"deleted": deleted, "failures": failures}, nil
}

// --- Version-history bloat ---

type VersionBloat struct {
	Ref         string `json:"ref"` // API path to the item
	Name        string `json:"name"`
	Path        string `json:"path"`
	Versions    int    `json:"versions"`
	CurrentSize int64  `json:"currentSize"`
	Reclaimable int64  `json:"reclaimable"` // total size of non-current versions
}

// FindVersionBloat finds files whose version history wastes space. minVersions
// filters out files with few versions; maxFiles caps how many files we probe
// (version lookups are one call each).
func (cl *CleanupService) FindVersionBloat(ownerType, ownerID string, minVersions, maxFiles int) ([]VersionBloat, error) {
	if minVersions < 2 {
		minVersions = 2
	}
	if maxFiles <= 0 {
		maxFiles = 3000
	}
	cl.cancel.Store(false)
	files, err := cl.walkFiles(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	// Probe the largest files first — that's where version bloat matters.
	sort.Slice(files, func(i, j int) bool { return files[i].size > files[j].size })
	if len(files) > maxFiles {
		files = files[:maxFiles]
	}

	out := make([]VersionBloat, 0)
	for idx, f := range files {
		if cl.cancel.Load() {
			cl.emit("Canceled", idx, len(files))
			break
		}
		if idx%25 == 0 {
			cl.emit("Checking versions", idx, len(files))
		}
		var vr struct {
			Value []struct {
				Size int64 `json:"size"`
			} `json:"value"`
		}
		ref := f.base + "/items/" + url.PathEscape(f.id)
		if err := c.Get(cl.s.Ctx(), ref+"/versions", url.Values{"$select": {"id,size"}}, &vr); err != nil {
			continue
		}
		if len(vr.Value) < minVersions {
			continue
		}
		var total int64
		for _, v := range vr.Value {
			total += v.Size
		}
		reclaimable := total - f.size
		if reclaimable <= 0 {
			continue
		}
		out = append(out, VersionBloat{
			Ref: ref, Name: f.name, Path: f.path, Versions: len(vr.Value),
			CurrentSize: f.size, Reclaimable: reclaimable,
		})
	}
	cl.emit("Done", len(files), len(files))
	sort.Slice(out, func(i, j int) bool { return out[i].Reclaimable > out[j].Reclaimable })
	return out, nil
}

// TrimVersions deletes old versions of an item, keeping the newest `keep`.
// Destructive: confirm == "TRIM".
func (cl *CleanupService) TrimVersions(itemRef string, keep int, confirm string) (map[string]any, error) {
	if err := cl.s.GuardWrite(); err != nil {
		return nil, err
	}
	if confirm != "TRIM" {
		return nil, errors.New("type TRIM to confirm version trimming")
	}
	if keep < 1 {
		keep = 1
	}
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	var vr struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	if err := c.Get(cl.s.Ctx(), itemRef+"/versions", url.Values{"$select": {"id"}}, &vr); err != nil {
		return nil, err
	}
	// Versions come newest-first; keep the first `keep`, delete the rest.
	removed := 0
	failures := map[string]string{}
	for i, v := range vr.Value {
		if i < keep || v.ID == "current" {
			continue
		}
		if e := c.Delete(cl.s.Ctx(), itemRef+"/versions/"+url.PathEscape(v.ID)); e != nil {
			failures[v.ID] = e.Error()
			continue
		}
		removed++
	}
	cl.s.Record("cleanup.trimVersions", itemRef, "removed="+itoa(removed), nil)
	return map[string]any{"removed": removed, "failures": failures}, nil
}
