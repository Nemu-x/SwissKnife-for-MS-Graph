package services

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"sync"
	"time"
	"unicode"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// SnapshotService — "what changed in my tenant's security configuration since
// last week?". A snapshot collects a fixed set of read-only configuration
// objects into one JSON document saved under <configDir>/snapshots; a diff
// compares two snapshots object by object and field by field.
//
// Everything here is read-only against Graph. Reads use Policy.Read.All (CA,
// named locations, authorization / authentication-methods policies),
// RoleManagement.Read.Directory (directory roles + members),
// Directory.Read.All (SKUs, domains, groups, service principal count) and
// Application.Read.All (app registrations with credential expiry). A section
// the app may not read is recorded as skipped — the snapshot still succeeds.
type SnapshotService struct {
	// lastTaken makes TakenAt strictly increasing across back-to-back snapshots.
	mu        sync.Mutex
	lastTaken time.Time
	s         *session.Session
}

func NewSnapshotService(s *session.Session) *SnapshotService { return &SnapshotService{s: s} }

// SnapshotSection describes one collected configuration area.
type SnapshotSection struct {
	Name    string `json:"name"`
	Count   int    `json:"count"`
	Skipped bool   `json:"skipped,omitempty"`
	Error   string `json:"error,omitempty"`
}

// SnapshotMeta is what the UI lists: identity, when, what was captured.
type SnapshotMeta struct {
	ID       string            `json:"id"`
	Name     string            `json:"name"`
	TakenAt  time.Time         `json:"takenAt"`
	Tenant   string            `json:"tenant,omitempty"` // connection profile name
	Sections []SnapshotSection `json:"sections"`
}

// snapshotFile is the on-disk document: meta plus one array of normalized
// objects per section (single-object sections are one-element arrays).
type snapshotFile struct {
	Meta     SnapshotMeta                `json:"meta"`
	Sections map[string][]map[string]any `json:"sections"`
}

// FieldChange is one leaf-level difference inside a changed object.
type FieldChange struct {
	Path   string `json:"path"`
	Before any    `json:"before"`
	After  any    `json:"after"`
}

// ObjectChange is one object that was added, removed or changed in a section.
// Added/removed carry the object itself; changed carry the field changes.
type ObjectChange struct {
	Key     string         `json:"key"`
	Label   string         `json:"label"`
	Changes []FieldChange  `json:"changes,omitempty"`
	Object  map[string]any `json:"object,omitempty"`
}

// SectionDiff is the per-section result of comparing two snapshots.
type SectionDiff struct {
	Name      string         `json:"name"`
	Skipped   bool           `json:"skipped,omitempty"` // missing or skipped on either side — not comparable
	Note      string         `json:"note,omitempty"`
	Added     []ObjectChange `json:"added"`
	Removed   []ObjectChange `json:"removed"`
	Changed   []ObjectChange `json:"changed"`
	Unchanged int            `json:"unchanged"`
}

// SnapshotDiff compares snapshot A (older) with B (newer).
type SnapshotDiff struct {
	A        SnapshotMeta  `json:"a"`
	B        SnapshotMeta  `json:"b"`
	Sections []SectionDiff `json:"sections"`
	Added    int           `json:"added"`
	Removed  int           `json:"removed"`
	Changed  int           `json:"changed"`
}

// snapshotListCap bounds the paged collections (groups, applications) so a
// huge tenant cannot turn a snapshot into a multi-minute crawl (ADR-003).
const snapshotListCap = 5000

type sectionCollector struct {
	name    string
	collect func(ctx context.Context, c *graphapi.Client) ([]map[string]any, error)
}

// snapshotSections is the fixed capture set, in the order they are collected
// (also the order sections are reported in a diff).
var snapshotSections = []sectionCollector{
	{"conditionalAccessPolicies", listSection("/identity/conditionalAccess/policies", nil, 0)},
	{"namedLocations", listSection("/identity/conditionalAccess/namedLocations", nil, 0)},
	{"directoryRoles", collectDirectoryRoles},
	{"authorizationPolicy", getSection("/policies/authorizationPolicy")},
	{"authenticationMethodsPolicy", getSection("/policies/authenticationMethodsPolicy")},
	{"subscribedSkus", listSection("/subscribedSkus", url.Values{
		"$select": {"id,skuId,skuPartNumber,capabilityStatus,consumedUnits,prepaidUnits"},
	}, 0)},
	{"domains", listSection("/domains", nil, 0)},
	{"groups", listSection("/groups", url.Values{
		"$select": {"id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule"},
		"$top":    {"999"},
	}, snapshotListCap)},
	{"applications", listSection("/applications", url.Values{
		"$select": {"id,appId,displayName,passwordCredentials,keyCredentials"},
		"$top":    {"999"},
	}, snapshotListCap)},
	{"servicePrincipals", collectServicePrincipalCount},
}

func listSection(path string, params url.Values, maxItems int) func(context.Context, *graphapi.Client) ([]map[string]any, error) {
	return func(ctx context.Context, c *graphapi.Client) ([]map[string]any, error) {
		raw, err := c.ListAll(ctx, path, params, maxItems)
		if err != nil {
			return nil, err
		}
		return decodeObjects(raw)
	}
}

func getSection(path string) func(context.Context, *graphapi.Client) ([]map[string]any, error) {
	return func(ctx context.Context, c *graphapi.Client) ([]map[string]any, error) {
		var obj map[string]any
		if err := c.Get(ctx, path, nil, &obj); err != nil {
			return nil, err
		}
		return []map[string]any{obj}, nil
	}
}

// collectDirectoryRoles captures every activated directory role together with
// its members — "who is Global Administrator" is the drift people care about.
func collectDirectoryRoles(ctx context.Context, c *graphapi.Client) ([]map[string]any, error) {
	raw, err := c.ListAll(ctx, "/directoryRoles", nil, 0)
	if err != nil {
		return nil, err
	}
	roles, err := decodeObjects(raw)
	if err != nil {
		return nil, err
	}
	params := url.Values{"$select": {"id,userPrincipalName,displayName"}}
	for _, role := range roles {
		id, _ := role["id"].(string)
		if id == "" {
			continue
		}
		rawMembers, err := c.ListAll(ctx, "/directoryRoles/"+url.PathEscape(id)+"/members", params, 0)
		if err != nil {
			return nil, err
		}
		members, err := decodeObjects(rawMembers)
		if err != nil {
			return nil, err
		}
		list := make([]any, 0, len(members))
		for _, m := range members {
			list = append(list, m)
		}
		role["members"] = list
	}
	return roles, nil
}

// collectServicePrincipalCount records only how many enterprise apps exist —
// the full list is large and its consents have their own review tile.
func collectServicePrincipalCount(ctx context.Context, c *graphapi.Client) ([]map[string]any, error) {
	n, err := c.Count(ctx, "/servicePrincipals")
	if err != nil {
		return nil, err
	}
	return []map[string]any{{"id": "servicePrincipals", "count": n}}, nil
}

func decodeObjects(raw []json.RawMessage) ([]map[string]any, error) {
	out := make([]map[string]any, 0, len(raw))
	for i, r := range raw {
		var m map[string]any
		if err := json.Unmarshal(r, &m); err != nil {
			return nil, fmt.Errorf("decode item %d: %w", i, err)
		}
		out = append(out, m)
	}
	return out, nil
}

// Take collects every section (best-effort: a section the app cannot read is
// marked skipped with its error) and saves the snapshot to disk.
func (x *SnapshotService) Take(name string) (*SnapshotMeta, error) {
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	dir, err := x.dir()
	if err != nil {
		return nil, err
	}
	ctx := x.s.Ctx()
	name = strings.TrimSpace(name)
	if name == "" {
		name = "snapshot"
	}

	doc := snapshotFile{Sections: map[string][]map[string]any{}}
	total := len(snapshotSections)
	skipped := 0
	var firstErr error
	for i, sec := range snapshotSections {
		emitEvent(ctx, "snapshot:progress", map[string]any{"section": sec.name, "done": i, "total": total})
		info := SnapshotSection{Name: sec.name}
		objs, err := sec.collect(ctx, c)
		switch {
		case err != nil && ctx.Err() != nil:
			return nil, ctx.Err() // shutdown / cancel: nothing to save
		case err != nil:
			info.Skipped = true
			info.Error = wrapOpErr(err).Error()
			skipped++
			if firstErr == nil {
				firstErr = err
			}
		default:
			for j := range objs {
				objs[j] = normalizeValue(objs[j]).(map[string]any)
			}
			sortObjects(objs)
			info.Count = len(objs)
			doc.Sections[sec.name] = objs
		}
		doc.Meta.Sections = append(doc.Meta.Sections, info)
	}
	emitEvent(ctx, "snapshot:progress", map[string]any{"section": "", "done": total, "total": total})
	if skipped == total {
		// Nothing readable at all (expired token, no permissions): a snapshot
		// of nothing would only produce a misleading "everything removed" diff.
		return nil, fmt.Errorf("every section failed — first error: %w", firstErr)
	}

	now := x.nextStamp()
	doc.Meta.Name = name
	doc.Meta.TakenAt = now
	doc.Meta.Tenant = x.s.ProfileName()
	doc.Meta.ID = uniqueSnapshotID(dir, now, name)
	b, err := json.Marshal(doc)
	if err == nil {
		err = os.WriteFile(filepath.Join(dir, doc.Meta.ID+".json"), b, 0o600)
	}
	x.s.Record("snapshot.take", doc.Meta.ID, fmt.Sprintf("%d sections, %d skipped", total, skipped), err)
	if err != nil {
		return nil, err
	}
	return &doc.Meta, nil
}

// List returns the saved snapshots, newest first.
// nextStamp returns a strictly increasing timestamp. On Windows the clock can
// tick every few milliseconds, so two snapshots taken back to back would
// otherwise share TakenAt and their order in List/DiffLatest would be undefined.
func (x *SnapshotService) nextStamp() time.Time {
	x.mu.Lock()
	defer x.mu.Unlock()
	now := time.Now()
	if !now.After(x.lastTaken) {
		now = x.lastTaken.Add(time.Microsecond)
	}
	x.lastTaken = now
	return now
}

func (x *SnapshotService) List() ([]SnapshotMeta, error) {
	dir, err := x.dir()
	if err != nil {
		return nil, err
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, err
	}
	out := []SnapshotMeta{}
	for _, e := range entries {
		if e.IsDir() || !strings.HasSuffix(e.Name(), ".json") {
			continue
		}
		doc, err := x.load(strings.TrimSuffix(e.Name(), ".json"))
		if err != nil {
			continue // a corrupt file must not hide the healthy ones
		}
		out = append(out, doc.Meta)
	}
	sort.Slice(out, func(i, j int) bool {
		if !out[i].TakenAt.Equal(out[j].TakenAt) {
			return out[i].TakenAt.After(out[j].TakenAt)
		}
		return out[i].ID > out[j].ID
	})
	return out, nil
}

// Get returns the whole snapshot document (meta + sections) as raw JSON.
func (x *SnapshotService) Get(id string) (json.RawMessage, error) {
	path, err := x.path(id)
	if err != nil {
		return nil, err
	}
	b, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	if !json.Valid(b) {
		return nil, fmt.Errorf("snapshot %s is not valid JSON", id)
	}
	return json.RawMessage(b), nil
}

// Delete removes a saved snapshot file.
func (x *SnapshotService) Delete(id string) error {
	path, err := x.path(id)
	if err != nil {
		return err
	}
	err = os.Remove(path)
	x.s.Record("snapshot.delete", id, "", err)
	return err
}

// Diff compares snapshot a (older) with b (newer).
func (x *SnapshotService) Diff(aID, bID string) (*SnapshotDiff, error) {
	a, err := x.load(aID)
	if err != nil {
		return nil, err
	}
	b, err := x.load(bID)
	if err != nil {
		return nil, err
	}
	return diffSnapshots(a, b), nil
}

// DiffLatest compares the two newest snapshots (older → newer).
func (x *SnapshotService) DiffLatest() (*SnapshotDiff, error) {
	list, err := x.List()
	if err != nil {
		return nil, err
	}
	if len(list) < 2 {
		return nil, errors.New("at least two snapshots are needed to compare")
	}
	return x.Diff(list[1].ID, list[0].ID)
}

// --- storage helpers ---

func (x *SnapshotService) dir() (string, error) {
	base := x.s.ConfigDir()
	if base == "" {
		return "", errors.New("config directory is not set")
	}
	d := filepath.Join(base, "snapshots")
	if err := os.MkdirAll(d, 0o700); err != nil {
		return "", err
	}
	return d, nil
}

// path resolves a snapshot id to its file, refusing anything that could
// escape the snapshots directory (ids come from the frontend).
func (x *SnapshotService) path(id string) (string, error) {
	if !validSnapshotID(id) {
		return "", fmt.Errorf("invalid snapshot id %q", id)
	}
	dir, err := x.dir()
	if err != nil {
		return "", err
	}
	return filepath.Join(dir, id+".json"), nil
}

func validSnapshotID(id string) bool {
	return id != "" && !strings.HasPrefix(id, ".") && !strings.Contains(id, "..") &&
		!strings.ContainsAny(id, `/\:`) && filepath.Base(id) == id
}

func (x *SnapshotService) load(id string) (*snapshotFile, error) {
	path, err := x.path(id)
	if err != nil {
		return nil, err
	}
	b, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var doc snapshotFile
	if err := json.Unmarshal(b, &doc); err != nil {
		return nil, fmt.Errorf("snapshot %s: %w", id, err)
	}
	if doc.Meta.ID == "" {
		doc.Meta.ID = id
	}
	if doc.Sections == nil {
		doc.Sections = map[string][]map[string]any{}
	}
	return &doc, nil
}

// uniqueSnapshotID builds "<timestamp>-<slug>" and appends -2, -3… when two
// snapshots land in the same second.
func uniqueSnapshotID(dir string, now time.Time, name string) string {
	base := now.Format("20060102-150405") + "-" + slugify(name)
	id := base
	for n := 2; ; n++ {
		if _, err := os.Stat(filepath.Join(dir, id+".json")); os.IsNotExist(err) {
			return id
		}
		id = fmt.Sprintf("%s-%d", base, n)
	}
}

func slugify(name string) string {
	var b strings.Builder
	dash := false
	for _, r := range strings.ToLower(name) {
		if unicode.IsLetter(r) || unicode.IsDigit(r) {
			b.WriteRune(r)
			dash = false
		} else if !dash && b.Len() > 0 {
			b.WriteByte('-')
			dash = true
		}
	}
	s := strings.TrimRight(b.String(), "-")
	if runes := []rune(s); len(runes) > 40 {
		s = strings.TrimRight(string(runes[:40]), "-")
	}
	if s == "" {
		return "snapshot"
	}
	return s
}

// --- normalization ---

// normalizeValue strips volatile OData annotations (keeping @odata.type, which
// tells an IP named location from a country one) and sorts arrays so two
// snapshots of an unchanged tenant compare equal byte for byte.
func normalizeValue(v any) any {
	switch t := v.(type) {
	case map[string]any:
		for k, val := range t {
			if strings.Contains(k, "@odata") && k != "@odata.type" {
				delete(t, k)
				continue
			}
			t[k] = normalizeValue(val)
		}
		return t
	case []any:
		for i := range t {
			t[i] = normalizeValue(t[i])
		}
		sortAny(t)
		return t
	default:
		return v
	}
}

// keyOf identifies an object inside an array: id, then keyId (credentials),
// then displayName, else its canonical JSON.
func keyOf(o map[string]any) string {
	for _, k := range []string{"id", "keyId", "displayName"} {
		if s, ok := o[k].(string); ok && s != "" {
			return s
		}
	}
	b, _ := json.Marshal(o) // Go sorts map keys: canonical
	return string(b)
}

// labelOf is the human name for a diff row.
func labelOf(o map[string]any) string {
	for _, k := range []string{"displayName", "userPrincipalName", "skuPartNumber", "id", "keyId"} {
		if s, ok := o[k].(string); ok && s != "" {
			return s
		}
	}
	return keyOf(o)
}

func allObjects(arr []any) bool {
	for _, v := range arr {
		if _, ok := v.(map[string]any); !ok {
			return false
		}
	}
	return true
}

func allStrings(arr []any) bool {
	for _, v := range arr {
		if _, ok := v.(string); !ok {
			return false
		}
	}
	return true
}

func sortAny(arr []any) {
	switch {
	case len(arr) < 2:
	case allObjects(arr):
		sort.SliceStable(arr, func(i, j int) bool {
			return keyOf(arr[i].(map[string]any)) < keyOf(arr[j].(map[string]any))
		})
	case allStrings(arr):
		sort.SliceStable(arr, func(i, j int) bool { return arr[i].(string) < arr[j].(string) })
	}
}

func sortObjects(objs []map[string]any) {
	sort.SliceStable(objs, func(i, j int) bool { return keyOf(objs[i]) < keyOf(objs[j]) })
}

// --- diff ---

// diffIgnore lists fields kept in the snapshot for reference but excluded from
// the comparison: they move without any configuration change.
var diffIgnore = map[string]bool{
	"modifiedDateTime":     true,
	"lastModifiedDateTime": true,
}

func diffSnapshots(a, b *snapshotFile) *SnapshotDiff {
	out := &SnapshotDiff{A: a.Meta, B: b.Meta}
	infoA := sectionIndex(a.Meta)
	infoB := sectionIndex(b.Meta)

	// Known sections in capture order, then anything else (older files).
	var names []string
	seen := map[string]bool{}
	for _, s := range snapshotSections {
		names = append(names, s.name)
		seen[s.name] = true
	}
	var extra []string
	for n := range infoA {
		if !seen[n] {
			seen[n] = true
			extra = append(extra, n)
		}
	}
	for n := range infoB {
		if !seen[n] {
			seen[n] = true
			extra = append(extra, n)
		}
	}
	sort.Strings(extra)
	names = append(names, extra...)

	for _, name := range names {
		ia, okA := infoA[name]
		ib, okB := infoB[name]
		if !okA && !okB {
			continue
		}
		sd := SectionDiff{Name: name, Added: []ObjectChange{}, Removed: []ObjectChange{}, Changed: []ObjectChange{}}
		switch {
		case !okA || !okB:
			sd.Skipped, sd.Note = true, "missing in one snapshot"
		case ia.Skipped || ib.Skipped:
			sd.Skipped = true
			sd.Note = ia.Error
			if sd.Note == "" {
				sd.Note = ib.Error
			}
		default:
			diffObjects(a.Sections[name], b.Sections[name], &sd)
		}
		out.Added += len(sd.Added)
		out.Removed += len(sd.Removed)
		out.Changed += len(sd.Changed)
		out.Sections = append(out.Sections, sd)
	}
	return out
}

func sectionIndex(m SnapshotMeta) map[string]SnapshotSection {
	idx := make(map[string]SnapshotSection, len(m.Sections))
	for _, s := range m.Sections {
		idx[s.Name] = s
	}
	return idx
}

// indexObjects keys a section by object identity, disambiguating collisions
// (two objects with the same displayName and no id) with a #n suffix.
func indexObjects(objs []map[string]any) ([]string, map[string]map[string]any) {
	keys := make([]string, 0, len(objs))
	idx := make(map[string]map[string]any, len(objs))
	for _, o := range objs {
		k := keyOf(o)
		for n := 2; ; n++ {
			if _, dup := idx[k]; !dup {
				break
			}
			k = fmt.Sprintf("%s#%d", keyOf(o), n)
		}
		keys = append(keys, k)
		idx[k] = o
	}
	return keys, idx
}

func diffObjects(objsA, objsB []map[string]any, sd *SectionDiff) {
	keysA, idxA := indexObjects(objsA)
	keysB, idxB := indexObjects(objsB)
	for _, k := range keysA {
		if _, ok := idxB[k]; !ok {
			sd.Removed = append(sd.Removed, ObjectChange{Key: k, Label: labelOf(idxA[k]), Object: idxA[k]})
		}
	}
	for _, k := range keysB {
		oa, ok := idxA[k]
		if !ok {
			sd.Added = append(sd.Added, ObjectChange{Key: k, Label: labelOf(idxB[k]), Object: idxB[k]})
			continue
		}
		if changes := diffJSON("", oa, idxB[k]); len(changes) > 0 {
			sd.Changed = append(sd.Changed, ObjectChange{Key: k, Label: labelOf(idxB[k]), Changes: changes})
		} else {
			sd.Unchanged++
		}
	}
}

// diffJSON walks two decoded JSON values and lists the leaf differences with
// dotted paths ("conditions.users.includeUsers", "members[<id>]"). Arrays of
// identifiable objects are matched by key; other arrays compare as a whole.
func diffJSON(path string, a, b any) []FieldChange {
	switch av := a.(type) {
	case map[string]any:
		bv, ok := b.(map[string]any)
		if !ok {
			return []FieldChange{{Path: path, Before: a, After: b}}
		}
		keys := make([]string, 0, len(av)+len(bv))
		seen := map[string]bool{}
		for k := range av {
			seen[k] = true
			keys = append(keys, k)
		}
		for k := range bv {
			if !seen[k] {
				keys = append(keys, k)
			}
		}
		sort.Strings(keys)
		var out []FieldChange
		for _, k := range keys {
			if diffIgnore[k] {
				continue
			}
			p := k
			if path != "" {
				p = path + "." + k
			}
			ca, inA := av[k]
			cb, inB := bv[k]
			switch {
			case !inA:
				out = append(out, FieldChange{Path: p, Before: nil, After: cb})
			case !inB:
				out = append(out, FieldChange{Path: p, Before: ca, After: nil})
			default:
				out = append(out, diffJSON(p, ca, cb)...)
			}
		}
		return out
	case []any:
		bv, ok := b.([]any)
		if !ok {
			return []FieldChange{{Path: path, Before: a, After: b}}
		}
		if (len(av) > 0 || len(bv) > 0) && allObjects(av) && allObjects(bv) {
			var out []FieldChange
			objsA := make([]map[string]any, len(av))
			for i := range av {
				objsA[i] = av[i].(map[string]any)
			}
			objsB := make([]map[string]any, len(bv))
			for i := range bv {
				objsB[i] = bv[i].(map[string]any)
			}
			keysA, idxA := indexObjects(objsA)
			keysB, idxB := indexObjects(objsB)
			for _, k := range keysA {
				if _, ok := idxB[k]; !ok {
					out = append(out, FieldChange{Path: path + "[" + k + "]", Before: idxA[k], After: nil})
				}
			}
			for _, k := range keysB {
				oa, ok := idxA[k]
				if !ok {
					out = append(out, FieldChange{Path: path + "[" + k + "]", Before: nil, After: idxB[k]})
					continue
				}
				out = append(out, diffJSON(path+"["+k+"]", oa, idxB[k])...)
			}
			return out
		}
		if !reflect.DeepEqual(a, b) {
			return []FieldChange{{Path: path, Before: a, After: b}}
		}
		return nil
	default:
		if !reflect.DeepEqual(a, b) {
			return []FieldChange{{Path: path, Before: a, After: b}}
		}
		return nil
	}
}
