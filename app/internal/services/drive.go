package services

import (
	"context"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// DriveService serves both OneDrive and SharePoint drives.
// ownerType: "user" (a user OneDrive) | "site" (a SharePoint site drive).
type DriveService struct {
	s *session.Session
}

// CancelTransfer cancels the live transfer operation: its context aborts
// in-flight HTTP and the copy returns its partial result with Canceled = true.
func (d *DriveService) CancelTransfer() { d.s.Ops.CancelKind(ops.KindTransfer) }

func NewDriveService(s *session.Session) *DriveService { return &DriveService{s: s} }

func drivePath(ownerType, ownerID string) (string, error) {
	switch ownerType {
	case "user":
		return "/users/" + url.PathEscape(ownerID) + "/drive", nil
	case "site":
		return "/sites/" + url.PathEscape(ownerID) + "/drive", nil
	}
	return "", errors.New("ownerType must be 'user' or 'site'")
}

// Sites searches SharePoint sites.
func (d *DriveService) Sites(search string) ([]json.RawMessage, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{"$top": {"50"}}
	if search != "" {
		params.Set("search", search)
	}
	return c.ListAll(d.s.Ctx(), "/sites", params, 200)
}

// SiteUsage is a SharePoint site with its default document library storage used,
// so the operator can spot the biggest sites before scanning.
type SiteUsage struct {
	ID     string `json:"id"`
	Name   string `json:"name"`
	WebURL string `json:"webUrl"`
	Used   int64  `json:"used"` // -1 when the size could not be read
}

// SitesWithUsage lists sites and fetches each one's storage used (default drive
// quota), then sorts largest-first. Sizes are read concurrently to stay fast.
func (d *DriveService) SitesWithUsage(search string) ([]SiteUsage, error) {
	raw, err := d.Sites(search)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	type siteLite struct {
		ID          string `json:"id"`
		Name        string `json:"name"`
		DisplayName string `json:"displayName"`
		WebURL      string `json:"webUrl"`
	}
	out := make([]SiteUsage, 0, len(raw))
	var mu sync.Mutex
	var wg sync.WaitGroup
	sem := make(chan struct{}, 8) // bound concurrency so we don't hammer Graph
	for _, r := range raw {
		var s siteLite
		if json.Unmarshal(r, &s) != nil || s.ID == "" {
			continue
		}
		// This picker targets shared spaces; skip personal OneDrive sites.
		if strings.Contains(strings.ToLower(s.WebURL), "/personal/") {
			continue
		}
		name := s.DisplayName
		if name == "" {
			name = s.Name
		}
		if name == "" {
			name = s.ID
		}
		wg.Add(1)
		sem <- struct{}{}
		go func(id, name, webURL string) {
			defer wg.Done()
			defer func() { <-sem }()
			used := int64(-1)
			var drive struct {
				Quota struct {
					Used int64 `json:"used"`
				} `json:"quota"`
			}
			if err := c.Get(d.s.Ctx(), "/sites/"+url.PathEscape(id)+"/drive", url.Values{"$select": {"quota"}}, &drive); err == nil {
				used = drive.Quota.Used
			}
			mu.Lock()
			out = append(out, SiteUsage{ID: id, Name: name, WebURL: webURL, Used: used})
			mu.Unlock()
		}(s.ID, name, s.WebURL)
	}
	wg.Wait()
	sort.Slice(out, func(a, b int) bool { return out[a].Used > out[b].Used })
	return out, nil
}

func (d *DriveService) ListRoot(ownerType, ownerID string) ([]json.RawMessage, error) {
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(d.s.Ctx(), base+"/root/children", nil, 0)
}

func (d *DriveService) Children(ownerType, ownerID, itemID string) ([]json.RawMessage, error) {
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(d.s.Ctx(), base+"/items/"+url.PathEscape(itemID)+"/children", nil, 0)
}

func (d *DriveService) Search(ownerType, ownerID, query string) ([]json.RawMessage, error) {
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(d.s.Ctx(), base+"/root/search(q='"+url.PathEscape(escapeODataLiteral(query))+"')", nil, 200)
}

func (d *DriveService) emitProgress(op *ops.Operation, name string, done, total int64) {
	emitOp(d.s.Ctx(), op, "transfer:progress", map[string]any{
		"name": name, "done": done, "total": total,
	})
}

// emitFile reports one item's outcome plus running counters so the UI can keep
// a live log even when the operator leaves and returns to the page.
func (d *DriveService) emitFile(op *ops.Operation, name, status, reason string, copied int) {
	emitOp(d.s.Ctx(), op, "transfer:file", map[string]any{
		"name": name, "status": status, "reason": reason, "copied": copied,
	})
}

// emitOverall reports whole-transfer progress (bytes processed vs. the scanned
// total), letting the UI render an overall percentage. totalBytes may be 0 when
// no preview was available; the UI then falls back to file counters.
func (d *DriveService) emitOverall(op *ops.Operation, doneBytes, totalBytes int64, doneFiles, totalFiles int) {
	emitOp(d.s.Ctx(), op, "transfer:overall", map[string]any{
		"doneBytes": doneBytes, "totalBytes": totalBytes, "files": doneFiles, "totalFiles": totalFiles,
	})
}

// Download shows a save dialog and streams to disk with progress.
func (d *DriveService) Download(ownerType, ownerID, itemID, suggestedName string) (string, error) {
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return "", err
	}
	c, err := d.s.Client()
	if err != nil {
		return "", err
	}
	local, err := wrt.SaveFileDialog(d.s.Ctx(), wrt.SaveDialogOptions{
		DefaultFilename: suggestedName,
		Title:           "Save file",
	})
	if err != nil || local == "" {
		return "", err // dialog cancellation is not an error
	}
	err = c.DownloadItem(d.s.Ctx(), base+"/items/"+url.PathEscape(itemID), local,
		func(done, total int64) { d.emitProgress(nil, suggestedName, done, total) })
	d.s.Record("drive.download", ownerType+":"+ownerID, "item="+itemID, err)
	if err != nil {
		return "", err
	}
	return local, nil
}

// Upload shows a file picker and uploads (small files via PUT, large via chunked session).
// remoteFolder is the drive folder path ("" = root); the name comes from the file.
func (d *DriveService) Upload(ownerType, ownerID, remoteFolder string) (json.RawMessage, error) {
	if err := d.s.GuardWrite(); err != nil {
		return nil, err
	}
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	local, err := wrt.OpenFileDialog(d.s.Ctx(), wrt.OpenDialogOptions{Title: "Choose file to upload"})
	if err != nil || local == "" {
		return nil, err
	}
	name := filepath.Base(local)
	remote := strings.Trim(remoteFolder, "/")
	if remote != "" {
		remote += "/"
	}
	uploadRoot := base + "/root:/" + escapeDrivePath(remote+name)
	out, err := c.UploadFile(d.s.Ctx(), uploadRoot, local,
		func(done, total int64) { d.emitProgress(nil, name, done, total) })
	d.s.Record("drive.upload", ownerType+":"+ownerID, "path="+remote+name, err)
	return out, err
}

func (d *DriveService) Delete(ownerType, ownerID, itemID, confirm string) error {
	if err := d.s.GuardDestructive(itemID, confirm); err != nil {
		return err
	}
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return err
	}
	c, err := d.s.Client()
	if err != nil {
		return err
	}
	err = c.Delete(d.s.Ctx(), base+"/items/"+url.PathEscape(itemID))
	d.s.Record("drive.delete", ownerType+":"+ownerID, "item="+itemID, err)
	return err
}

func (d *DriveService) CreateLink(ownerType, ownerID, itemID, linkType, scope string) (json.RawMessage, error) {
	if err := d.s.GuardWrite(); err != nil {
		return nil, err
	}
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	if linkType == "" {
		linkType = "view"
	}
	if scope == "" {
		scope = "organization"
	}
	var out json.RawMessage
	err = c.Post(d.s.Ctx(), base+"/items/"+url.PathEscape(itemID)+"/createLink",
		map[string]any{"type": linkType, "scope": scope}, &out)
	d.s.Record("drive.createLink", ownerType+":"+ownerID, "item="+itemID+" type="+linkType+" scope="+scope, err)
	return out, err
}

// CopyResult is the outcome of a OneDrive-to-OneDrive copy.
type CopyResult struct {
	Copied   []string          `json:"copied"`
	Skipped  map[string]string `json:"skipped"` // name -> reason
	Failed   map[string]string `json:"failed"`  // name -> error
	Canceled bool              `json:"canceled"`
}

// CopyPreview summarizes what an offboarding copy would transfer.
type CopyPreview struct {
	Files      int   `json:"files"`
	Folders    int   `json:"folders"`
	TotalBytes int64 `json:"totalBytes"`
}

// DriveQuota is a user's OneDrive storage quota.
type DriveQuota struct {
	Total     int64 `json:"total"`
	Used      int64 `json:"used"`
	Remaining int64 `json:"remaining"`
}

// Quota returns the user's OneDrive storage quota.
func (d *DriveService) Quota(user string) (*DriveQuota, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	var drive struct {
		Quota DriveQuota `json:"quota"`
	}
	if err := c.Get(d.s.Ctx(), "/users/"+url.PathEscape(user)+"/drive", url.Values{"$select": {"quota"}}, &drive); err != nil {
		return nil, err
	}
	return &drive.Quota, nil
}

// PickTarget returns the first target user whose OneDrive has at least
// neededBytes free — used to rotate offboarding backups across a pool.
func (d *DriveService) PickTarget(targets []string, neededBytes int64) (string, error) {
	for _, tgt := range targets {
		tgt = strings.TrimSpace(tgt)
		if tgt == "" {
			continue
		}
		q, err := d.Quota(tgt)
		if err != nil {
			continue // skip targets we can't read
		}
		if q.Remaining >= neededBytes {
			return tgt, nil
		}
	}
	return "", errors.New("no target in the pool has enough free space")
}

// OffboardingPreview walks the source user's OneDrive (read-only) and reports
// the file/folder count and total size, so the operator can review before copying.
func (d *DriveService) OffboardingPreview(sourceUser string) (*CopyPreview, error) {
	prev := &CopyPreview{}
	var walk func(itemID string) error
	count := func(items []json.RawMessage) {
		for _, raw := range items {
			var it struct {
				ID     string          `json:"id"`
				Size   int64           `json:"size"`
				Folder json.RawMessage `json:"folder"`
			}
			if json.Unmarshal(raw, &it) != nil {
				continue
			}
			if it.Folder != nil {
				prev.Folders++
				_ = walk(it.ID)
				continue
			}
			prev.Files++
			prev.TotalBytes += it.Size
		}
	}
	walk = func(itemID string) error {
		items, err := d.Children("user", sourceUser, itemID)
		if err != nil {
			return err
		}
		count(items)
		return nil
	}
	rootItems, err := d.ListRoot("user", sourceUser)
	if err != nil {
		return nil, err
	}
	count(rootItems)
	return prev, nil
}

// CopyBetweenUsers recursively copies OneDrive source to target via the OS temp
// directory (no hardcoded /tmp — fixes the legacy bug). destFolder is an optional
// subfolder in the target drive (e.g. "Backups/alice"); "" copies into the root.
func (d *DriveService) CopyBetweenUsers(sourceUser, targetUser, destFolder string, overwrite bool) (*CopyResult, error) {
	res, err := d.copyBetweenUsersCtx(d.s.Ctx(), sourceUser, targetUser, destFolder, overwrite, nil)
	return res, wrapOpErr(err)
}

// copyBetweenUsersCtx is CopyBetweenUsers with an explicit parent context (a
// playbook passes its operation context so cancelling the playbook cancels the
// copy) and an optional pre-computed preview. It first attempts a cloud-side
// copy (Graph copies items inside the service — bytes never transit the
// operator's machine, folders copy recursively in one operation). If that
// cannot even start, it falls back to the legacy download/upload path through
// the local temp directory.
func (d *DriveService) copyBetweenUsersCtx(parent context.Context, sourceUser, targetUser, destFolder string, overwrite bool, prev *CopyPreview) (res *CopyResult, err error) {
	if err := d.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	op, err := d.s.Ops.Start(parent, ops.KindTransfer)
	if err != nil {
		return nil, err
	}
	defer d.s.Ops.Finish(op)
	emitOp(d.s.Ctx(), op, "op:start", map[string]any{"target": sourceUser + " → " + targetUser})
	if j := d.s.Journal; j != nil {
		j.Begin(op.ID, map[string]any{
			"kind": "transfer", "target": sourceUser + " → " + targetUser,
			"source": sourceUser, "dest": targetUser, "folder": destFolder, "overwrite": overwrite,
		})
		defer func() { j.End(op.ID, copySummary(res, err)) }()
	}

	if res, sErr := d.copyServerSide(op, c, sourceUser, targetUser, destFolder, overwrite); sErr == nil {
		return res, nil
	} else {
		d.emitFile(op, "(server-side copy unavailable — using local copy)", "skipped", sErr.Error(), 0)
	}
	if op.Canceled() {
		return &CopyResult{Skipped: map[string]string{}, Failed: map[string]string{}, Canceled: true}, nil
	}
	if prev == nil {
		prev, _ = d.OffboardingPreview(sourceUser) // progress totals only; copy proceeds without them
	}
	var totalBytes int64
	var totalFiles int
	if prev != nil {
		totalBytes, totalFiles = prev.TotalBytes, prev.Files
	}
	var doneBytes int64
	doneFiles := 0
	d.emitOverall(op, 0, totalBytes, 0, totalFiles)
	tmp, err := os.MkdirTemp("", "swissknife-copy-*")
	if err != nil {
		return nil, err
	}
	defer os.RemoveAll(tmp)

	dest := strings.Trim(destFolder, "/")
	res = &CopyResult{Skipped: map[string]string{}, Failed: map[string]string{}}

	// names in the target root, for the overwrite check (top level only)
	tgtNames := map[string]bool{}
	if !overwrite && dest == "" {
		items, err := d.ListRoot("user", targetUser)
		if err != nil {
			return nil, err
		}
		for _, raw := range items {
			var it struct {
				Name string `json:"name"`
			}
			if json.Unmarshal(raw, &it) == nil {
				tgtNames[it.Name] = true
			}
		}
	}

	var walk func(itemID, rel string) error
	copyItems := func(items []json.RawMessage, rel string) {
		for _, raw := range items {
			if op.Canceled() {
				return
			}
			var it struct {
				ID     string          `json:"id"`
				Name   string          `json:"name"`
				Size   int64           `json:"size"`
				Folder json.RawMessage `json:"folder"`
			}
			if json.Unmarshal(raw, &it) != nil || it.Name == "" {
				continue
			}
			relPath := it.Name
			if rel != "" {
				relPath = rel + "/" + it.Name
			}
			if it.Folder != nil {
				if err := walk(it.ID, relPath); err != nil {
					res.Failed[relPath+"/"] = err.Error()
					d.emitFile(op, relPath+"/", "failed", err.Error(), len(res.Copied))
				}
				continue
			}
			// Every processed file (copied, skipped or failed) advances the
			// overall byte counter, so the percentage always reaches 100.
			account := func() {
				doneBytes += it.Size
				doneFiles++
				d.emitOverall(op, doneBytes, totalBytes, doneFiles, totalFiles)
			}
			if rel == "" && tgtNames[it.Name] {
				res.Skipped[relPath] = "exists in target"
				d.emitFile(op, relPath, "skipped", "exists in target", len(res.Copied))
				account()
				continue
			}
			local := filepath.Join(tmp, filepath.FromSlash(relPath))
			src := "/users/" + url.PathEscape(sourceUser) + "/drive/items/" + url.PathEscape(it.ID)
			if err := c.DownloadItem(op.Ctx, src, local, nil); err != nil {
				res.Failed[relPath] = err.Error()
				d.emitFile(op, relPath, "failed", err.Error(), len(res.Copied))
				account()
				continue
			}
			remotePath := relPath
			if dest != "" {
				remotePath = dest + "/" + relPath
			}
			dst := "/users/" + url.PathEscape(targetUser) + "/drive/root:/" + escapeDrivePath(remotePath)
			if _, err := c.UploadFile(op.Ctx, dst, local, func(done, total int64) {
				d.emitProgress(op, relPath, done, total)
			}); err != nil {
				res.Failed[relPath] = err.Error()
				d.emitFile(op, relPath, "failed", err.Error(), len(res.Copied))
				account()
				continue
			}
			res.Copied = append(res.Copied, relPath)
			d.emitFile(op, relPath, "copied", "", len(res.Copied))
			account()
			os.Remove(local)
		}
	}

	walk = func(itemID, rel string) error {
		if op.Canceled() {
			return nil
		}
		items, err := d.Children("user", sourceUser, itemID)
		if err != nil {
			return err
		}
		copyItems(items, rel)
		return nil
	}

	rootItems, err := d.ListRoot("user", sourceUser)
	if err != nil {
		return nil, err
	}
	copyItems(rootItems, "")
	res.Canceled = op.Canceled()

	d.s.Record("drive.copyBetweenUsers", sourceUser+" -> "+targetUser,
		"dest="+dest+" copied="+itoa(len(res.Copied))+" skipped="+itoa(len(res.Skipped))+" failed="+itoa(len(res.Failed)), nil)
	return res, nil
}

// errCopyCanceled aborts the server-side polling loop on operator cancel.
var errCopyCanceled = errors.New("canceled")

// copySummary is the journal terminal record for a copy run.
func copySummary(res *CopyResult, err error) map[string]any {
	out := map[string]any{}
	if err != nil {
		out["error"] = err.Error()
	}
	if res != nil {
		out["copied"] = len(res.Copied)
		out["skipped"] = len(res.Skipped)
		out["failed"] = len(res.Failed)
		out["canceled"] = res.Canceled
	}
	return out
}

// journalItem records one server-side copy item's lifecycle for resume.
func (d *DriveService) journalItem(opID string, data map[string]any) {
	if j := d.s.Journal; j != nil {
		j.Event(opID, "item", data)
	}
}

// copyServerSide copies the source drive's top-level items into the target via
// Graph's async driveItem copy — the copy runs entirely inside the cloud.
// An error return means nothing was copied yet and the caller may fall back;
// per-item failures after that are recorded in the result instead.
func (d *DriveService) copyServerSide(op *ops.Operation, c *graphapi.Client, sourceUser, targetUser, destFolder string, overwrite bool) (*CopyResult, error) {
	ctx := op.Ctx
	var drv struct {
		ID string `json:"id"`
	}
	if err := c.Get(ctx, "/users/"+url.PathEscape(targetUser)+"/drive", url.Values{"$select": {"id"}}, &drv); err != nil {
		return nil, err
	}
	parentID, err := d.ensureFolder(c, targetUser, destFolder)
	if err != nil {
		return nil, err
	}
	rootItems, err := d.ListRoot("user", sourceUser)
	if err != nil {
		return nil, err
	}
	type srcItem struct {
		ID     string          `json:"id"`
		Name   string          `json:"name"`
		Size   int64           `json:"size"` // for folders Graph reports the total content size
		Folder json.RawMessage `json:"folder"`
	}
	items := make([]srcItem, 0, len(rootItems))
	var totalBytes int64
	for _, raw := range rootItems {
		var it srcItem
		if json.Unmarshal(raw, &it) != nil || it.Name == "" {
			continue
		}
		items = append(items, it)
		totalBytes += it.Size
	}

	res := &CopyResult{Skipped: map[string]string{}, Failed: map[string]string{}}
	behavior := "fail" // backup semantics: never touch what already exists
	if overwrite {
		behavior = "replace"
	}
	// Journal the copy plan so an interrupted run can be resumed after a
	// restart: remaining items are re-issued, in-flight monitors re-polled.
	if j := d.s.Journal; j != nil {
		plan := make([]map[string]any, 0, len(items))
		for _, it := range items {
			plan = append(plan, map[string]any{"id": it.ID, "name": it.Name, "size": it.Size, "folder": it.Folder != nil})
		}
		j.Event(op.ID, "plan", map[string]any{"items": plan, "driveId": drv.ID, "parentId": parentID, "behavior": behavior, "sourceUser": sourceUser})
	}
	var doneBytes int64
	d.emitOverall(op, 0, totalBytes, 0, len(items))
	for i, it := range items {
		if op.Canceled() {
			res.Canceled = true
			break
		}
		name := it.Name
		if it.Folder != nil {
			name += "/"
		}
		loc, postErr := c.PostForLocation(ctx,
			"/users/"+url.PathEscape(sourceUser)+"/drive/items/"+url.PathEscape(it.ID)+"/copy",
			url.Values{"@microsoft.graph.conflictBehavior": {behavior}},
			map[string]any{"parentReference": map[string]any{"driveId": drv.ID, "id": parentID}})
		if postErr != nil {
			var ge *graphapi.GraphError
			if errors.As(postErr, &ge) && ge.Code == "nameAlreadyExists" {
				res.Skipped[name] = "exists in target"
				d.emitFile(op, name, "skipped", "exists in target", len(res.Copied))
				d.journalItem(op.ID, map[string]any{"id": it.ID, "status": "skipped"})
			} else if len(res.Copied) == 0 && len(res.Failed) == 0 {
				return nil, postErr // nothing copied yet — safe to fall back to local copy
			} else {
				res.Failed[name] = postErr.Error()
				d.emitFile(op, name, "failed", postErr.Error(), len(res.Copied))
				d.journalItem(op.ID, map[string]any{"id": it.ID, "status": "failed", "error": postErr.Error()})
			}
		} else {
			d.journalItem(op.ID, map[string]any{"id": it.ID, "status": "inflight", "monitor": loc})
			if waitErr := d.waitForCopy(op, loc, name, it.Size, doneBytes, totalBytes, len(res.Copied), len(items)); waitErr != nil {
				if errors.Is(waitErr, errCopyCanceled) {
					res.Canceled = true
					res.Skipped[name] = "cancel requested — the in-flight cloud copy may still finish"
					d.emitFile(op, name, "skipped", "canceled", len(res.Copied))
					break
				}
				res.Failed[name] = waitErr.Error()
				d.emitFile(op, name, "failed", waitErr.Error(), len(res.Copied))
				d.journalItem(op.ID, map[string]any{"id": it.ID, "status": "failed", "error": waitErr.Error()})
			} else {
				res.Copied = append(res.Copied, name)
				d.emitFile(op, name, "copied", "", len(res.Copied))
				d.journalItem(op.ID, map[string]any{"id": it.ID, "status": "copied"})
			}
		}
		doneBytes += it.Size
		d.emitOverall(op, doneBytes, totalBytes, i+1, len(items))
	}

	d.s.Record("drive.copyServerSide", sourceUser+" -> "+targetUser,
		"dest="+strings.Trim(destFolder, "/")+" copied="+itoa(len(res.Copied))+" skipped="+itoa(len(res.Skipped))+" failed="+itoa(len(res.Failed)), nil)
	return res, nil
}

// ResumeCopy continues an interrupted server-side copy from its journal:
// already-completed items are counted, journaled in-flight monitors re-polled,
// pending/failed items re-issued (conflictBehavior=fail turns anything that
// actually finished earlier into a "exists in target" skip). The continuation
// appends to the same journal file, which finally gets its terminal record.
func (d *DriveService) ResumeCopy(runID string) (res *CopyResult, err error) {
	if err := d.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	j := d.s.Journal
	if j == nil {
		return nil, errors.New("run journal unavailable")
	}
	run, err := j.Get(runID)
	if err != nil {
		return nil, err
	}
	if run.Kind != "transfer" || run.EndedAt != nil || !run.Interrupted {
		return nil, errors.New("this run is not an interrupted transfer")
	}

	// Reconstruct the plan and each item's last journaled status.
	type planItem struct {
		id, name, monitor, status string
		size                      int64
	}
	var driveID, parentID, behavior, sourceUser string
	items := map[string]*planItem{}
	order := []string{}
	for _, ev := range run.Events {
		switch ev.Type {
		case "plan":
			driveID, _ = ev.Data["driveId"].(string)
			parentID, _ = ev.Data["parentId"].(string)
			behavior, _ = ev.Data["behavior"].(string)
			sourceUser, _ = ev.Data["sourceUser"].(string)
			raw, _ := ev.Data["items"].([]any)
			for _, ri := range raw {
				m, _ := ri.(map[string]any)
				id, _ := m["id"].(string)
				if id == "" {
					continue
				}
				name, _ := m["name"].(string)
				size, _ := m["size"].(float64)
				items[id] = &planItem{id: id, name: name, size: int64(size), status: "pending"}
				order = append(order, id)
			}
		case "item":
			id, _ := ev.Data["id"].(string)
			it := items[id]
			if it == nil {
				continue
			}
			it.status, _ = ev.Data["status"].(string)
			if m, ok := ev.Data["monitor"].(string); ok {
				it.monitor = m
			}
		}
	}
	if len(order) == 0 || driveID == "" || sourceUser == "" {
		return nil, errors.New("this run has no resumable copy plan")
	}
	if behavior == "" {
		behavior = "fail"
	}

	op, err := d.s.Ops.Start(d.s.Ctx(), ops.KindTransfer)
	if err != nil {
		return nil, err
	}
	defer d.s.Ops.Finish(op)
	emitOp(d.s.Ctx(), op, "op:start", map[string]any{"target": run.Target})
	j.Event(runID, "log", map[string]any{"text": "resumed"})

	res = &CopyResult{Skipped: map[string]string{}, Failed: map[string]string{}}
	defer func() { j.End(runID, copySummary(res, err)) }()

	var totalBytes, doneBytes int64
	for _, id := range order {
		totalBytes += items[id].size
	}
	total := len(order)
	d.emitOverall(op, 0, totalBytes, 0, total)

	// issue re-posts one item's copy and waits for it.
	issue := func(it *planItem) {
		loc, postErr := c.PostForLocation(op.Ctx,
			"/users/"+url.PathEscape(sourceUser)+"/drive/items/"+url.PathEscape(it.id)+"/copy",
			url.Values{"@microsoft.graph.conflictBehavior": {behavior}},
			map[string]any{"parentReference": map[string]any{"driveId": driveID, "id": parentID}})
		if postErr != nil {
			var ge *graphapi.GraphError
			if errors.As(postErr, &ge) && ge.Code == "nameAlreadyExists" {
				res.Skipped[it.name] = "exists in target"
				d.emitFile(op, it.name, "skipped", "exists in target", len(res.Copied))
				d.journalItem(runID, map[string]any{"id": it.id, "status": "skipped"})
			} else {
				res.Failed[it.name] = postErr.Error()
				d.emitFile(op, it.name, "failed", postErr.Error(), len(res.Copied))
				d.journalItem(runID, map[string]any{"id": it.id, "status": "failed", "error": postErr.Error()})
			}
			return
		}
		d.journalItem(runID, map[string]any{"id": it.id, "status": "inflight", "monitor": loc})
		if waitErr := d.waitForCopy(op, loc, it.name, it.size, doneBytes, totalBytes, len(res.Copied), total); waitErr != nil {
			if errors.Is(waitErr, errCopyCanceled) {
				res.Canceled = true
				res.Skipped[it.name] = "cancel requested — the in-flight cloud copy may still finish"
				d.emitFile(op, it.name, "skipped", "canceled", len(res.Copied))
				return
			}
			res.Failed[it.name] = waitErr.Error()
			d.emitFile(op, it.name, "failed", waitErr.Error(), len(res.Copied))
			d.journalItem(runID, map[string]any{"id": it.id, "status": "failed", "error": waitErr.Error()})
			return
		}
		res.Copied = append(res.Copied, it.name)
		d.emitFile(op, it.name, "copied", "", len(res.Copied))
		d.journalItem(runID, map[string]any{"id": it.id, "status": "copied"})
	}

	for _, id := range order {
		it := items[id]
		if op.Canceled() {
			res.Canceled = true
			break
		}
		switch it.status {
		case "copied":
			res.Copied = append(res.Copied, it.name) // finished in the previous run
		case "skipped":
			res.Skipped[it.name] = "exists in target"
		case "inflight":
			// The cloud kept copying while we were gone — re-poll its monitor;
			// if the monitor died with the old session, fall back to re-issue.
			if it.monitor != "" {
				if waitErr := d.waitForCopy(op, it.monitor, it.name, it.size, doneBytes, totalBytes, len(res.Copied), total); waitErr == nil {
					res.Copied = append(res.Copied, it.name)
					d.emitFile(op, it.name, "copied", "", len(res.Copied))
					d.journalItem(runID, map[string]any{"id": it.id, "status": "copied"})
				} else if errors.Is(waitErr, errCopyCanceled) {
					res.Canceled = true
				} else {
					issue(it)
				}
			} else {
				issue(it)
			}
		default: // pending / failed / interrupted
			issue(it)
		}
		if res.Canceled {
			break
		}
		doneBytes += it.size
		d.emitOverall(op, doneBytes, totalBytes, len(res.Copied), total)
	}

	d.s.Record("drive.resumeCopy", run.Target,
		"run="+runID+" copied="+itoa(len(res.Copied))+" skipped="+itoa(len(res.Skipped))+" failed="+itoa(len(res.Failed)), nil)
	return res, nil
}

// ensureFolder creates destFolder (segment by segment) in the user's drive and
// returns its item id; "" resolves to the drive root.
func (d *DriveService) ensureFolder(c *graphapi.Client, user, destFolder string) (string, error) {
	ctx := d.s.Ctx()
	base := "/users/" + url.PathEscape(user) + "/drive"
	folder := strings.Trim(destFolder, "/")
	if folder == "" {
		var root struct {
			ID string `json:"id"`
		}
		if err := c.Get(ctx, base+"/root", url.Values{"$select": {"id"}}, &root); err != nil {
			return "", err
		}
		return root.ID, nil
	}
	built := ""
	for _, seg := range strings.Split(folder, "/") {
		parent := base + "/root/children"
		if built != "" {
			parent = base + "/root:/" + escapeDrivePath(built) + ":/children"
		}
		// Best-effort create; an existing folder answers nameAlreadyExists.
		_ = c.Post(ctx, parent, map[string]any{
			"name": seg, "folder": map[string]any{}, "@microsoft.graph.conflictBehavior": "fail",
		}, nil)
		if built == "" {
			built = seg
		} else {
			built += "/" + seg
		}
	}
	var f struct {
		ID string `json:"id"`
	}
	if err := c.Get(ctx, base+"/root:/"+escapeDrivePath(built), url.Values{"$select": {"id"}}, &f); err != nil {
		return "", err
	}
	return f.ID, nil
}

// monitorClient polls pre-authenticated Graph monitor URLs: no Authorization
// header, and redirects (303 on success) must not be followed.
var monitorClient = &http.Client{
	Timeout:       30 * time.Second,
	CheckRedirect: func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse },
}

// waitForCopy polls a Graph async-operation monitor URL until the copy ends,
// emitting fractional overall progress along the way. Transient poll errors are
// tolerated; a stuck operation gives up after maxCopyWait.
func (d *DriveService) waitForCopy(op *ops.Operation, monitorURL, name string, size, doneBytes, totalBytes int64, copied, totalItems int) error {
	if monitorURL == "" {
		return nil // some copies complete synchronously (204/201 without a monitor)
	}
	const maxCopyWait = 2 * time.Hour
	const maxPollErrors = 5
	ctx := op.Ctx
	deadline := time.Now().Add(maxCopyWait)
	pollErrors := 0
	for {
		if op.Canceled() {
			return errCopyCanceled
		}
		if time.Now().After(deadline) {
			return errors.New("server-side copy did not finish within " + maxCopyWait.String())
		}
		req, err := http.NewRequestWithContext(ctx, http.MethodGet, monitorURL, nil)
		if err != nil {
			return err
		}
		resp, err := monitorClient.Do(req)
		if err != nil {
			// Transient network hiccups must not fail a copy that keeps
			// running server-side — tolerate a few before giving up.
			pollErrors++
			if pollErrors > maxPollErrors {
				return err
			}
		} else {
			pollErrors = 0
			body, _ := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
			resp.Body.Close()
			if resp.StatusCode == http.StatusSeeOther { // redirect to the created item
				return nil
			}
			var st struct {
				Status             string  `json:"status"`
				PercentageComplete float64 `json:"percentageComplete"`
				StatusDescription  string  `json:"statusDescription"`
			}
			_ = json.Unmarshal(body, &st)
			switch st.Status {
			case "completed":
				return nil
			case "failed", "deleteFailed":
				msg := st.StatusDescription
				if msg == "" {
					msg = "server-side copy " + st.Status
				}
				return errors.New(msg)
			}
			// notStarted / waiting / inProgress / updating — keep polling.
			d.emitOverall(op, doneBytes+int64(float64(size)*st.PercentageComplete/100), totalBytes, copied, totalItems)
			d.emitProgress(op, name, int64(st.PercentageComplete), 100)
		}
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-time.After(1500 * time.Millisecond):
		}
	}
}

func itoa(n int) string { return strconv.Itoa(n) }

// escapeDrivePath encodes drive path segments (root:/{path}:) while preserving "/".
func escapeDrivePath(p string) string {
	segs := strings.Split(p, "/")
	for i, s := range segs {
		segs[i] = url.PathEscape(s)
	}
	return strings.Join(segs, "/")
}
