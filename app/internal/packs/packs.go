// Package packs loads community action packs (ADR-008 D13): a folder with a
// manifest.yaml and PowerShell scripts that adds actions to the catalog. A
// pack is loaded only when it is signed by a trusted minisign key or the
// operator trusted its exact contents (SHA-256); any change needs trust again.
package packs

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"unicode/utf8"

	"github.com/jedisct1/go-minisign"
	"gopkg.in/yaml.v3"
)

// Files that carry the signature; they are not part of the digest.
const (
	DigestFile    = "pack.digest"
	SignatureFile = "pack.digest.minisig"
	ManifestFile  = "manifest.yaml"
)

// BuiltinKey is the project's minisign public key (minisign.pub): packs it
// signs are trusted without asking.
const BuiltinKey = "RWSG5GttbIKtpCqGS1IxYmyO07AbVOztAbwaWSi22mlNsVzJTXHRLH+9"

// Manifest describes a pack.
type Manifest struct {
	Name        string      `yaml:"name" json:"name"`
	Version     string      `yaml:"version" json:"version"`
	Author      string      `yaml:"author" json:"author"`
	Description string      `yaml:"description" json:"description"`
	Actions     []ActionDef `yaml:"actions" json:"actions"`
}

// ActionDef is one action of a pack.
type ActionDef struct {
	ID           string            `yaml:"id" json:"id"`
	Page         string            `yaml:"page" json:"page"`
	Danger       string            `yaml:"danger" json:"danger"`
	Module       string            `yaml:"module" json:"module"` // exo | teams
	Script       string            `yaml:"script" json:"script"`
	Label        map[string]string `yaml:"label" json:"label"`
	Hint         map[string]string `yaml:"hint" json:"hint"`
	ConfirmField string            `yaml:"confirmField" json:"confirmField"`
	Columns      []string          `yaml:"columns" json:"columns"`
	Fields       []FieldDef        `yaml:"fields" json:"fields"`
}

// FieldDef is one input of a pack action.
type FieldDef struct {
	Name     string            `yaml:"name" json:"name"`
	Kind     string            `yaml:"kind" json:"kind"` // text | choice | user | group
	Required bool              `yaml:"required" json:"required"`
	Options  []string          `yaml:"options" json:"options"`
	Default  string            `yaml:"default" json:"default"`
	Label    map[string]string `yaml:"label" json:"label"`
}

// Status is a pack's trust state.
type Status string

const (
	Signed    Status = "signed"    // signature verified against a trusted key
	Trusted   Status = "trusted"   // unsigned, contents pinned by the operator
	Untrusted Status = "untrusted" // never trusted
	Changed   Status = "changed"   // pinned contents differ from the folder
	Disabled  Status = "disabled"  // turned off by the operator (even if signed)
	Invalid   Status = "invalid"   // cannot be loaded (manifest, signature, files)
)

// Pack is a loaded folder.
type Pack struct {
	Dir      string   `json:"dir"`
	Manifest Manifest `json:"manifest"`
	Digest   string   `json:"digest"`
	Status   Status   `json:"status"`
	Signer   string   `json:"signer,omitempty"` // key id of a verified signature
	Error    string   `json:"error,omitempty"`
	// Scripts maps a script path to its text (BOM removed), for running.
	Scripts map[string]string `json:"-"`
}

// Usable reports whether the pack's actions may run.
func (p Pack) Usable() bool { return p.Status == Signed || p.Status == Trusted }

// Trust is the operator's trust store (packs.json).
type Trust struct {
	Keys     []string          `json:"keys"`               // extra minisign public keys
	Pinned   map[string]string `json:"pinned"`             // pack name → trusted digest
	Disabled []string          `json:"disabled,omitempty"` // pack names turned off
}

var (
	nameRe  = regexp.MustCompile(`^[a-z0-9][a-z0-9-]{1,39}$`)
	idRe    = regexp.MustCompile(`^[a-z][A-Za-z0-9]{1,39}$`)
	fieldRe = regexp.MustCompile(`^[a-z][A-Za-z0-9]{0,30}$`)
)

// Pages pack actions may appear on.
var Pages = map[string]bool{"users": true, "groups": true, "teams": true, "mail": true, "security": true, "reports": true, "intune": true}

// Digest hashes every file of the pack except the signature files: one line
// per file ("path\x00sha256\n", sorted, forward slashes), hashed again.
func Digest(dir string) (string, error) {
	d, _, err := readPack(dir)
	return d, err
}

// readPack reads every file once and returns the digest with the bytes it
// was computed from: the manifest and scripts are parsed from these same
// bytes, so what was digested (signed, pinned) is exactly what runs.
func readPack(dir string) (string, map[string][]byte, error) {
	var lines []string
	files := map[string][]byte{}
	err := filepath.WalkDir(dir, func(p string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		if d.Type()&fs.ModeSymlink != 0 {
			return fmt.Errorf("%s: symbolic links are not allowed in a pack", d.Name())
		}
		rel, err := filepath.Rel(dir, p)
		if err != nil {
			return err
		}
		rel = filepath.ToSlash(rel)
		if rel == DigestFile || rel == SignatureFile {
			return nil
		}
		b, err := os.ReadFile(p)
		if err != nil {
			return err
		}
		files[rel] = b
		sum := sha256.Sum256(b)
		lines = append(lines, rel+"\x00"+hex.EncodeToString(sum[:])+"\n")
		return nil
	})
	if err != nil {
		return "", nil, err
	}
	sort.Strings(lines)
	sum := sha256.Sum256([]byte(strings.Join(lines, "")))
	return hex.EncodeToString(sum[:]), files, nil
}

// Load reads every pack under root and decides its status.
func Load(root string, trust Trust) []Pack {
	entries, err := os.ReadDir(root)
	if err != nil {
		return nil
	}
	keys := trustedKeys(trust.Keys)
	var out []Pack
	count := map[string]int{}
	for _, e := range entries {
		if !e.IsDir() {
			continue
		}
		p := loadOne(filepath.Join(root, e.Name()), trust, keys)
		count[p.Manifest.Name]++
		out = append(out, p)
	}
	// Two folders claiming one name: neither loads, so a look-alike can
	// never stand in for the real pack.
	for i := range out {
		if out[i].Manifest.Name != "" && count[out[i].Manifest.Name] > 1 {
			out[i].Status, out[i].Error = Invalid, "another pack folder has the same name — remove one"
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Dir < out[j].Dir })
	return out
}

// signingKey is a trusted key and how it is shown.
type signingKey struct {
	key  minisign.PublicKey
	name string
}

func trustedKeys(user []string) map[[8]byte][]signingKey {
	keys := map[[8]byte][]signingKey{}
	builtin, _ := minisign.NewPublicKey(BuiltinKey)
	keys[builtin.KeyId] = []signingKey{{builtin, "SwissKnife project"}}
	for _, k := range user {
		pk, err := minisign.NewPublicKey(strings.TrimSpace(k))
		if err != nil || pk.KeyId == builtin.KeyId {
			continue // never shadows the project's key
		}
		keys[pk.KeyId] = append(keys[pk.KeyId], signingKey{pk, fmt.Sprintf("%X", pk.KeyId)})
	}
	return keys
}

func loadOne(dir string, trust Trust, keys map[[8]byte][]signingKey) Pack {
	p := Pack{Dir: dir, Status: Invalid}
	fail := func(format string, a ...any) Pack {
		p.Status, p.Error = Invalid, fmt.Sprintf(format, a...)
		return p
	}
	digest, files, err := readPack(dir)
	if err != nil {
		return fail("%v", err)
	}
	p.Digest = digest
	raw, ok := files[ManifestFile]
	if !ok {
		return fail("no %s", ManifestFile)
	}
	if err := yaml.Unmarshal(raw, &p.Manifest); err != nil {
		return fail("manifest: %v", err)
	}
	if p.Scripts, err = validate(files, &p.Manifest); err != nil {
		return fail("%v", err)
	}
	for _, n := range trust.Disabled {
		if n == p.Manifest.Name {
			p.Status = Disabled
			return p
		}
	}
	// A signature, when present, must verify: a pack is never downgraded
	// from "signed" to "pinned" silently.
	if _, err := os.Stat(filepath.Join(dir, SignatureFile)); err == nil {
		signer, err := verify(dir, p.Digest, keys)
		if err != nil {
			return fail("signature: %v", err)
		}
		p.Status, p.Signer = Signed, signer
		return p
	}
	switch pin, ok := trust.Pinned[p.Manifest.Name]; {
	case !ok:
		p.Status = Untrusted
	case pin == p.Digest:
		p.Status = Trusted
	default:
		p.Status = Changed
	}
	return p
}

func verify(dir, digest string, keys map[[8]byte][]signingKey) (string, error) {
	msg, err := os.ReadFile(filepath.Join(dir, DigestFile))
	if err != nil {
		return "", fmt.Errorf("%s is missing", DigestFile)
	}
	if strings.TrimSpace(string(msg)) != digest {
		return "", errors.New("the files do not match the signed digest")
	}
	sig, err := minisign.NewSignatureFromFile(filepath.Join(dir, SignatureFile))
	if err != nil {
		return "", err
	}
	candidates, ok := keys[sig.KeyId]
	if !ok {
		return "", errors.New("signed by an unknown key — add the author's public key to trust it")
	}
	for _, k := range candidates {
		if ok, err := k.key.Verify(msg, sig); err == nil && ok {
			return k.name, nil
		}
	}
	return "", errors.New("invalid signature")
}

// validate checks the manifest and returns the scripts it uses.
func validate(files map[string][]byte, m *Manifest) (map[string]string, error) {
	if !nameRe.MatchString(m.Name) {
		return nil, errors.New("name: lowercase letters, digits and dashes (2-40)")
	}
	if len(m.Actions) == 0 {
		return nil, errors.New("no actions")
	}
	scripts := map[string]string{}
	ids := map[string]bool{}
	for i, a := range m.Actions {
		where := fmt.Sprintf("action %d", i+1)
		if !idRe.MatchString(a.ID) || ids[a.ID] {
			return nil, fmt.Errorf("%s: id must be unique, a letter then letters/digits", where)
		}
		ids[a.ID] = true
		if !Pages[a.Page] {
			return nil, fmt.Errorf("%s: page %q is not one packs may use", a.ID, a.Page)
		}
		switch a.Danger {
		case "read", "write", "destructive":
		default:
			return nil, fmt.Errorf("%s: danger must be read, write or destructive", a.ID)
		}
		if a.Module != "exo" && a.Module != "teams" {
			return nil, fmt.Errorf("%s: module must be exo or teams", a.ID)
		}
		if a.Label["en"] == "" {
			return nil, fmt.Errorf("%s: an English label is required", a.ID)
		}
		names := map[string]bool{}
		for _, f := range a.Fields {
			if !fieldRe.MatchString(f.Name) || names[f.Name] {
				return nil, fmt.Errorf("%s: field names must be unique identifiers", a.ID)
			}
			names[f.Name] = true
			switch f.Kind {
			case "text", "user", "group":
			case "choice":
				if len(f.Options) == 0 {
					return nil, fmt.Errorf("%s.%s: a choice needs options", a.ID, f.Name)
				}
			default:
				return nil, fmt.Errorf("%s.%s: kind must be text, choice, user or group", a.ID, f.Name)
			}
		}
		if a.Danger == "destructive" && !names[a.ConfirmField] {
			return nil, fmt.Errorf("%s: a destructive action names its confirmField", a.ID)
		}
		// The script must be a .ps1 inside the pack folder.
		rel := filepath.ToSlash(filepath.Clean(a.Script))
		if !strings.HasSuffix(strings.ToLower(rel), ".ps1") || strings.HasPrefix(rel, "../") || rel == ".." || filepath.IsAbs(a.Script) || strings.Contains(rel, ":") {
			return nil, fmt.Errorf("%s: script must be a .ps1 file in the pack", a.ID)
		}
		if _, ok := scripts[rel]; !ok {
			b, ok := files[rel]
			if !ok {
				return nil, fmt.Errorf("%s: script %s not found", a.ID, rel)
			}
			// The host hashes the UTF-8 text it receives: anything else
			// would look trusted here and be refused there.
			if len(b) >= 2 && (b[0] == 0xFF && b[1] == 0xFE || b[0] == 0xFE && b[1] == 0xFF) || !utf8.Valid(b) {
				return nil, fmt.Errorf("%s: save the script as UTF-8", rel)
			}
			scripts[rel] = strings.TrimPrefix(string(b), string(rune(0xFEFF)))
		}
		m.Actions[i].Script = rel
	}
	return scripts, nil
}

// LoadTrust reads packs.json ("" or missing = empty store).
func LoadTrust(path string) Trust {
	t := Trust{Pinned: map[string]string{}}
	if b, err := os.ReadFile(path); err == nil {
		_ = json.Unmarshal(b, &t)
	}
	if t.Pinned == nil {
		t.Pinned = map[string]string{}
	}
	return t
}

// SaveTrust writes packs.json atomically.
func SaveTrust(path string, t Trust) error {
	b, err := json.MarshalIndent(t, "", "  ")
	if err != nil {
		return err
	}
	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, b, 0o600); err != nil {
		return err
	}
	return os.Rename(tmp, path)
}

// CheckKey validates a minisign public key (the base64 line).
func CheckKey(k string) error {
	_, err := minisign.NewPublicKey(strings.TrimSpace(k))
	return err
}
