package services

import (
	"encoding/json"
	"errors"
	"net/url"
	"sort"

	"swissknife-app/internal/session"
)

// CleanupService finds and removes duplicate files to reclaim OneDrive/SharePoint
// storage (e.g. the classic "30 copies of the same video").
type CleanupService struct {
	s *session.Session
}

func NewCleanupService(s *session.Session) *CleanupService { return &CleanupService{s: s} }

// DupItem is one file within a duplicate group.
type DupItem struct {
	ID   string `json:"id"`
	Name string `json:"name"`
	Path string `json:"path"`
}

// DupGroup is a set of files that are byte-identical (same hash, or same name+size).
type DupGroup struct {
	Name   string    `json:"name"`
	Size   int64     `json:"size"`
	Count  int       `json:"count"`
	Wasted int64     `json:"wasted"` // (count-1) * size — space freed if only one is kept
	Items  []DupItem `json:"items"`
}

// FindDuplicates walks the drive and groups byte-identical files.
func (cl *CleanupService) FindDuplicates(ownerType, ownerID string) ([]DupGroup, error) {
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}

	type fileRec struct {
		id, name, path string
		size           int64
	}
	groups := map[string][]fileRec{}

	var walk func(itemsPath, rel string) error
	walk = func(itemsPath, rel string) error {
		items, err := c.ListAll(cl.s.Ctx(), itemsPath, nil, 0)
		if err != nil {
			return err
		}
		for _, raw := range items {
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
				_ = walk(base+"/items/"+url.PathEscape(it.ID)+"/children", relPath)
				continue
			}
			// key: prefer a content hash, else fall back to name+size
			key := it.File.Hashes.QuickXor
			if key == "" {
				key = it.File.Hashes.Sha256
			}
			if key == "" {
				key = it.Name + "|" + itoa(int(it.Size))
			}
			groups[key] = append(groups[key], fileRec{it.ID, it.Name, relPath, it.Size})
		}
		return nil
	}
	if err := walk(base+"/root/children", ""); err != nil {
		return nil, err
	}

	out := make([]DupGroup, 0)
	for _, recs := range groups {
		if len(recs) < 2 {
			continue
		}
		g := DupGroup{Name: recs[0].name, Size: recs[0].size, Count: len(recs), Wasted: int64(len(recs)-1) * recs[0].size}
		for _, r := range recs {
			g.Items = append(g.Items, DupItem{ID: r.id, Name: r.name, Path: r.path})
		}
		out = append(out, g)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Wasted > out[j].Wasted })
	return out, nil
}

// DeleteItems removes the given drive items. Destructive: confirm must be "DELETE".
func (cl *CleanupService) DeleteItems(ownerType, ownerID string, itemIDs []string, confirm string) (map[string]any, error) {
	if err := cl.s.GuardWrite(); err != nil {
		return nil, err
	}
	if confirm != "DELETE" {
		return nil, errors.New(`type DELETE to confirm bulk deletion`)
	}
	base, err := drivePath(ownerType, ownerID)
	if err != nil {
		return nil, err
	}
	c, err := cl.s.Client()
	if err != nil {
		return nil, err
	}
	deleted := 0
	failures := map[string]string{}
	for _, id := range itemIDs {
		if e := c.Delete(cl.s.Ctx(), base+"/items/"+url.PathEscape(id)); e != nil {
			failures[id] = e.Error()
			continue
		}
		deleted++
	}
	cl.s.Record("cleanup.deleteItems", ownerType+":"+ownerID, "deleted="+itoa(deleted), nil)
	return map[string]any{"deleted": deleted, "failures": failures}, nil
}
