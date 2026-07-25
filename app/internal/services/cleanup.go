package services

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"sort"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// CleanupService reclaims OneDrive/SharePoint storage: duplicate files and
// version-history bloat. For SharePoint sites it scans every document library.
type CleanupService struct {
	s  *session.Session
	op *ops.Operation // the run in flight (single-flight makes this safe)
}

func NewCleanupService(s *session.Session) *CleanupService { return &CleanupService{s: s} }

// CancelScan cancels the live cleanup operation; the walkers abort between
// items and return whatever they found so far rather than an error.
func (cl *CleanupService) CancelScan() { cl.s.Ops.CancelKind(ops.KindCleanup) }

// stopped reports whether the scan in flight was cancelled. Reading the op
// context directly (instead of a bridged flag) cannot leak a stale cancel
// into the next scan.
func (cl *CleanupService) stopped() bool {
	op := cl.op
	return op != nil && op.Canceled()
}

// beginScan registers the cleanup operation for cancellation and progress.
func (cl *CleanupService) beginScan() (*ops.Operation, error) {
	op, err := cl.s.Ops.Start(cl.s.Ctx(), ops.KindCleanup)
	if err != nil {
		return nil, err
	}
	cl.op = op
	emitOp(cl.s.Ctx(), op, "op:start", nil)
	return op, nil
}

// endScan releases the operation slot and detaches it from the emit helpers.
func (cl *CleanupService) endScan(op *ops.Operation) {
	cl.op = nil
	cl.s.Ops.Finish(op)
}

// emit sends a live progress line to the UI ("cleanup:progress" event).
func (cl *CleanupService) emit(stage string, done, total int) {
	emitOp(cl.s.Ctx(), cl.op, "cleanup:progress", map[string]any{"stage": stage, "done": done, "total": total})
}

// emitLog appends a durable line to the UI console ("cleanup:log" event) so the
// operator can see exactly what was walked, not just a spinner.
func (cl *CleanupService) emitLog(text string) {
	emitOp(cl.s.Ctx(), cl.op, "cleanup:log", map[string]any{"text": text})
}

// humanSize formats a byte count as a short human-readable string.
func humanSize(b int64) string {
	const unit = 1024
	if b < unit {
		return fmt.Sprintf("%d B", b)
	}
	div, exp := int64(unit), 0
	for n := b / unit; n >= unit; n /= unit {
		div *= unit
		exp++
	}
	return fmt.Sprintf("%.1f %cB", float64(b)/float64(div), "KMGTPE"[exp])
}

// driveBases returns the API base path(s) for the owner's drive(s).
// A user has one drive; a site can have several document libraries.
func (cl *CleanupService) driveBases(ctx context.Context, ownerType, ownerID string) ([]string, error) {
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
		if err := c.Get(ctx, "/sites/"+url.PathEscape(ownerID)+"/drives", url.Values{"$select": {"id"}}, &resp); err == nil && len(resp.Value) > 0 {
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
func (cl *CleanupService) walkFiles(ctx context.Context, ownerType, ownerID string) ([]fileRec, error) {
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	bases, err := cl.driveBases(ctx, ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	if ownerType == "site" {
		cl.emitLog(fmt.Sprintf("Enumerating %d document librar%s…", len(bases), pluralY(len(bases))))
	}
	var files []fileRec
	var walk func(base, itemsPath, rel string) error
	walk = func(base, itemsPath, rel string) error {
		if cl.stopped() {
			return nil
		}
		items, err := c.ListAll(ctx, itemsPath, nil, 0)
		if err != nil {
			return err
		}
		for _, raw := range items {
			if cl.stopped() {
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
	for i, base := range bases {
		before := len(files)
		var bytes int64
		if err := walk(base, base+"/root/children", ""); err != nil {
			return nil, err
		}
		for _, f := range files[before:] {
			bytes += f.size
		}
		if len(bases) > 1 {
			cl.emitLog(fmt.Sprintf("Library %d/%d: %d files, %s", i+1, len(bases), len(files)-before, humanSize(bytes)))
		}
	}
	return files, nil
}

// pluralY returns "y" for 1 and "ies" otherwise, for "librar{y|ies}".
func pluralY(n int) string {
	if n == 1 {
		return "y"
	}
	return "ies"
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
	op, err := cl.beginScan()
	if err != nil {
		return nil, err
	}
	defer cl.endScan(op)
	files, err := cl.walkFiles(op.Ctx, ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	var volume int64
	for _, f := range files {
		volume += f.size
	}
	cl.emitLog(fmt.Sprintf("Volume: %d files, %s total", len(files), humanSize(volume)))
	if len(files) == 0 {
		cl.emitLog("No files found — the account may lack read access to this site's libraries (needs Sites.Read.All / Files.Read.All).")
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
	op, err := cl.beginScan()
	if err != nil {
		return nil, err
	}
	defer cl.endScan(op)
	files, err := cl.walkFiles(op.Ctx, ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	// Pre-scan summary: report the total volume so the operator can judge whether
	// a deep version scan is worthwhile before it runs.
	var volume int64
	for _, f := range files {
		volume += f.size
	}
	cl.emitLog(fmt.Sprintf("Volume: %d files, %s total", len(files), humanSize(volume)))
	if len(files) == 0 {
		cl.emitLog("No files found — the account may lack read access to this site's libraries (needs Sites.Read.All / Files.Read.All).")
		cl.emit("Done", 0, 0)
		return make([]VersionBloat, 0), nil
	}

	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	// Probe the largest files first — that's where version bloat matters.
	sort.Slice(files, func(i, j int) bool { return files[i].size > files[j].size })
	if len(files) > maxFiles {
		cl.emitLog(fmt.Sprintf("Probing the %d largest files (of %d) for version history…", maxFiles, len(files)))
		files = files[:maxFiles]
	} else {
		cl.emitLog(fmt.Sprintf("Probing %d files for version history…", len(files)))
	}

	out := make([]VersionBloat, 0)
	probeErrors := 0
	for idx, f := range files {
		if cl.stopped() {
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
		if err := c.Get(op.Ctx, ref+"/versions", url.Values{"$select": {"id,size"}}, &vr); err != nil {
			probeErrors++
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
	if probeErrors > 0 {
		cl.emitLog(fmt.Sprintf("%d file(s) could not be checked (version lookup failed).", probeErrors))
	}
	cl.emitLog(fmt.Sprintf("Found %d file(s) with reclaimable version history.", len(out)))
	sort.Slice(out, func(i, j int) bool { return out[i].Reclaimable > out[j].Reclaimable })
	return out, nil
}

// trimOne deletes old versions of a single item, keeping the newest `keep`.
// Versions come newest-first; keep the first `keep`, delete the rest.
func (cl *CleanupService) trimOne(ctx context.Context, c *graphapi.Client, itemRef string, keep int) (removed int, failures map[string]string, err error) {
	var vr struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	if err := c.Get(ctx, itemRef+"/versions", url.Values{"$select": {"id"}}, &vr); err != nil {
		return 0, nil, err
	}
	failures = map[string]string{}
	for i, v := range vr.Value {
		// A cancelled operation must stop deleting versions immediately, not
		// record each aborted call as a per-version failure.
		if cerr := ctx.Err(); cerr != nil {
			return removed, failures, cerr
		}
		if i < keep || v.ID == "current" {
			continue
		}
		if e := c.Delete(ctx, itemRef+"/versions/"+url.PathEscape(v.ID)); e != nil {
			failures[v.ID] = e.Error()
			continue
		}
		removed++
	}
	return removed, failures, nil
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
	removed, failures, err := cl.trimOne(cl.s.Ctx(), c, itemRef, keep)
	if err != nil {
		return nil, err
	}
	cl.s.Record("cleanup.trimVersions", itemRef, "removed="+itoa(removed), nil)
	return map[string]any{"removed": removed, "failures": failures}, nil
}

// TrimResult reports the outcome of trimming one item in a bulk run.
type TrimResult struct {
	Ref     string `json:"ref"`
	Removed int    `json:"removed"`
	Error   string `json:"error,omitempty"`
}

// TrimVersionsMany trims version history on many items in one confirmed run,
// streaming progress to the job console. Destructive: confirm == "TRIM".
// Per-item failures are reported, not fatal; CancelScan aborts between items.
func (cl *CleanupService) TrimVersionsMany(itemRefs []string, keep int, confirm string) ([]TrimResult, error) {
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
	op, err := cl.beginScan()
	if err != nil {
		return nil, err
	}
	defer cl.endScan(op)
	cl.emitLog(fmt.Sprintf("Trimming %d file(s), keeping the latest %d version(s)…", len(itemRefs), keep))
	out := make([]TrimResult, 0, len(itemRefs))
	totalRemoved, failed := 0, 0
	for i, ref := range itemRefs {
		if cl.stopped() {
			cl.emit("Canceled", i, len(itemRefs))
			cl.emitLog(fmt.Sprintf("Canceled after %d of %d file(s).", i, len(itemRefs)))
			break
		}
		cl.emit("Trimming versions", i, len(itemRefs))
		removed, failures, err := cl.trimOne(op.Ctx, c, ref, keep)
		r := TrimResult{Ref: ref, Removed: removed}
		if err != nil {
			r.Error = err.Error()
		} else if len(failures) > 0 {
			r.Error = fmt.Sprintf("%d version(s) could not be deleted", len(failures))
		}
		if r.Error != "" {
			failed++
			cl.emitLog(fmt.Sprintf("Failed: %s — %s", ref, r.Error))
		}
		totalRemoved += removed
		out = append(out, r)
	}
	cl.emit("Done", len(out), len(itemRefs))
	cl.emitLog(fmt.Sprintf("Trim done — %d version(s) removed across %d file(s), %d failure(s).", totalRemoved, len(out), failed))
	cl.s.Record("cleanup.trimVersionsMany", itoa(len(out))+" items", "removed="+itoa(totalRemoved), nil)
	return out, nil
}
