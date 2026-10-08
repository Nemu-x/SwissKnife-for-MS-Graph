package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"
)

// Configuration as code: a snapshot written out as one JSON file per object,
// in a folder the operator can put under version control. Ordering and
// formatting are stable so a commit shows exactly what changed.

// Secret-looking properties never leave the app, even where Graph returns
// them empty: exact names plus anything whose name suggests a secret.
var exportSecretKeys = map[string]bool{"key": true, "hint": true, "secretReferenceValueId": true}

func isSecretKey(k string) bool {
	if exportSecretKeys[k] {
		return true
	}
	l := strings.ToLower(k)
	for _, p := range []string{"secret", "password", "presharedkey", "passphrase", "privatekey", "certificatepassword"} {
		if strings.Contains(l, p) {
			return true
		}
	}
	return false
}

func stripSecrets(v any) any {
	switch t := v.(type) {
	case map[string]any:
		out := make(map[string]any, len(t))
		encrypted, _ := t["isEncrypted"].(bool) // Intune OMA-URI settings
		for k, x := range t {
			if isSecretKey(k) || (encrypted && k == "value") {
				continue
			}
			out[k] = stripSecrets(x)
		}
		return out
	case []any:
		out := make([]any, len(t))
		for i, x := range t {
			out[i] = stripSecrets(x)
		}
		return out
	default:
		return v
	}
}

// writeExport lays the snapshot out under root/<snapshot id>/ and returns
// that folder. An existing export of the same snapshot is replaced.
func writeExport(doc *snapshotFile, root string) (string, error) {
	if !validSnapshotID(doc.Meta.ID) {
		return "", errors.New("invalid snapshot id")
	}
	dir := filepath.Join(root, doc.Meta.ID)
	if err := os.RemoveAll(dir); err != nil {
		return "", err
	}
	write := func(path string, v any) error {
		b, err := json.MarshalIndent(v, "", "  ") // Go sorts map keys: stable
		if err != nil {
			return err
		}
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return err
		}
		return os.WriteFile(path, append(b, '\n'), 0o644)
	}
	if err := write(filepath.Join(dir, "snapshot.json"), doc.Meta); err != nil {
		return "", err
	}
	sections := make([]string, 0, len(doc.Sections))
	for name := range doc.Sections {
		sections = append(sections, name)
	}
	sort.Strings(sections)
	for _, name := range sections {
		used := map[string]int{}
		for _, obj := range doc.Sections[name] {
			base := slugify(labelOf(obj))
			if base == "" {
				base = "object"
			}
			// Same-named objects get the start of their key, then a counter.
			if used[base] > 0 {
				k := slugify(keyOf(obj))
				if len(k) > 8 {
					k = k[:8]
				}
				base = fmt.Sprintf("%s-%s", base, k)
			}
			used[base]++
			if used[base] > 1 {
				base = fmt.Sprintf("%s-%d", base, used[base])
			}
			if err := write(filepath.Join(dir, name, base+".json"), stripSecrets(obj)); err != nil {
				return "", err
			}
		}
	}
	return dir, nil
}

// Export writes a snapshot into a folder the operator picks and returns the
// folder it created.
func (x *SnapshotService) Export(id string) (string, error) {
	doc, err := x.load(id)
	if err != nil {
		return "", err
	}
	if x.s.Ctx().Value("events") == nil {
		return "", errors.New("folder dialogs need the desktop app")
	}
	root, err := wrt.OpenDirectoryDialog(x.s.Ctx(), wrt.OpenDialogOptions{Title: "Export snapshot"})
	if err != nil || root == "" {
		return "", err
	}
	dir, err := writeExport(doc, root)
	x.s.Record("snapshot.export", id, dir, err)
	return dir, err
}
