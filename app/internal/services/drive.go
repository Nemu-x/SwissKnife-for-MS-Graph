package services

import (
	"encoding/json"
	"errors"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/session"
)

// DriveService serves both OneDrive and SharePoint drives.
// ownerType: "user" (a user OneDrive) | "site" (a SharePoint site drive).
type DriveService struct {
	s *session.Session
}

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

func (d *DriveService) emitProgress(name string, done, total int64) {
	wrt.EventsEmit(d.s.Ctx(), "transfer:progress", map[string]any{
		"name": name, "done": done, "total": total,
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
		func(done, total int64) { d.emitProgress(suggestedName, done, total) })
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
		func(done, total int64) { d.emitProgress(name, done, total) })
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
	Copied  []string          `json:"copied"`
	Skipped map[string]string `json:"skipped"` // name -> reason
	Failed  map[string]string `json:"failed"`  // name -> error
}

// CopyBetweenUsers recursively copies OneDrive source to target via
// the OS temp directory (no hardcoded /tmp — fixes the legacy bug).
func (d *DriveService) CopyBetweenUsers(sourceUser, targetUser string, overwrite bool) (*CopyResult, error) {
	if err := d.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	tmp, err := os.MkdirTemp("", "swissknife-copy-*")
	if err != nil {
		return nil, err
	}
	defer os.RemoveAll(tmp)

	res := &CopyResult{Skipped: map[string]string{}, Failed: map[string]string{}}

	// names in the target root, for the overwrite check (top level only)
	tgtNames := map[string]bool{}
	if !overwrite {
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
			var it struct {
				ID     string          `json:"id"`
				Name   string          `json:"name"`
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
				}
				continue
			}
			if rel == "" && tgtNames[it.Name] {
				res.Skipped[relPath] = "exists in target"
				continue
			}
			local := filepath.Join(tmp, filepath.FromSlash(relPath))
			src := "/users/" + url.PathEscape(sourceUser) + "/drive/items/" + url.PathEscape(it.ID)
			if err := c.DownloadItem(d.s.Ctx(), src, local, nil); err != nil {
				res.Failed[relPath] = err.Error()
				continue
			}
			dst := "/users/" + url.PathEscape(targetUser) + "/drive/root:/" + escapeDrivePath(relPath)
			if _, err := c.UploadFile(d.s.Ctx(), dst, local, func(done, total int64) {
				d.emitProgress(relPath, done, total)
			}); err != nil {
				res.Failed[relPath] = err.Error()
				continue
			}
			res.Copied = append(res.Copied, relPath)
			os.Remove(local)
		}
	}

	walk = func(itemID, rel string) error {
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

	d.s.Record("drive.copyBetweenUsers", sourceUser+" -> "+targetUser,
		"copied="+itoa(len(res.Copied))+" skipped="+itoa(len(res.Skipped))+" failed="+itoa(len(res.Failed)), nil)
	return res, nil
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
