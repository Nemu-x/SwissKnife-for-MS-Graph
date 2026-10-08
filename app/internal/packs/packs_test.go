package packs

import (
	"crypto/ed25519"
	"crypto/rand"
	"encoding/base64"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const manifest = `name: contoso-tools
version: 1.0.0
author: Contoso IT
actions:
  - id: litigationHolds
    page: reports
    danger: read
    module: exo
    script: holds.ps1
    label: { en: Mailboxes on litigation hold, ru: Ящики на удержании }
    columns: [name, address]
`

func writePack(t *testing.T, root, dir, man string) string {
	t.Helper()
	d := filepath.Join(root, dir)
	if err := os.MkdirAll(d, 0o755); err != nil {
		t.Fatal(err)
	}
	_ = os.WriteFile(filepath.Join(d, ManifestFile), []byte(man), 0o644)
	_ = os.WriteFile(filepath.Join(d, "holds.ps1"), []byte(string(rune(0xFEFF))+"param($Mode, $Inputs)\nGet-Mailbox -ResultSize 10"), 0o644)
	return d
}

// sign writes pack.digest and a legacy-mode minisign signature made with a
// fresh key; it returns that key's public key line.
func sign(t *testing.T, dir string) string {
	t.Helper()
	pub, priv, _ := ed25519.GenerateKey(rand.Reader)
	keyID := []byte{1, 2, 3, 4, 5, 6, 7, 8}
	digest, err := Digest(dir)
	if err != nil {
		t.Fatal(err)
	}
	msg := []byte(digest + "\n")
	_ = os.WriteFile(filepath.Join(dir, DigestFile), msg, 0o644)
	sig := ed25519.Sign(priv, msg)
	trusted := "pack contoso-tools"
	global := ed25519.Sign(priv, append(append([]byte{}, sig...), []byte(trusted)...))
	b64 := base64.StdEncoding.EncodeToString
	text := "untrusted comment: test\n" + b64(append(append([]byte("Ed"), keyID...), sig...)) + "\n" +
		"trusted comment: " + trusted + "\n" + b64(global) + "\n"
	_ = os.WriteFile(filepath.Join(dir, SignatureFile), []byte(text), 0o644)
	return b64(append(append([]byte("Ed"), keyID...), pub...))
}

func TestPinnedPackBecomesChangedWhenItsFilesChange(t *testing.T) {
	root := t.TempDir()
	dir := writePack(t, root, "contoso", manifest)
	p := Load(root, Trust{})[0]
	if p.Status != Untrusted || p.Manifest.Actions[0].Label["ru"] == "" {
		t.Fatalf("new pack %+v", p)
	}
	if strings.HasPrefix(p.Scripts["holds.ps1"], string(rune(0xFEFF))) {
		t.Fatal("the BOM must be removed from scripts")
	}
	trust := Trust{Pinned: map[string]string{"contoso-tools": p.Digest}}
	if got := Load(root, trust)[0]; got.Status != Trusted || !got.Usable() {
		t.Fatalf("pinned pack %+v", got)
	}
	_ = os.WriteFile(filepath.Join(dir, "holds.ps1"), []byte("Remove-Mailbox -Identity *"), 0o644)
	if got := Load(root, trust)[0]; got.Status != Changed || got.Usable() {
		t.Fatalf("a changed pack must need trust again: %+v", got)
	}
}

func TestSignedPackVerifiesAgainstTrustedKeysOnly(t *testing.T) {
	root := t.TempDir()
	dir := writePack(t, root, "contoso", manifest)
	key := sign(t, dir)
	if got := Load(root, Trust{})[0]; got.Status != Invalid || !strings.Contains(got.Error, "unknown key") {
		t.Fatalf("unknown signer: %+v", got)
	}
	if err := CheckKey(key); err != nil {
		t.Fatal(err)
	}
	if got := Load(root, Trust{Keys: []string{key}})[0]; got.Status != Signed || got.Signer == "" {
		t.Fatalf("signed pack: %+v", got)
	}
	// Tampering after signing breaks it — never silently "unsigned".
	_ = os.WriteFile(filepath.Join(dir, "holds.ps1"), []byte("Remove-Mailbox -Identity *"), 0o644)
	if got := Load(root, Trust{Keys: []string{key}})[0]; got.Status != Invalid {
		t.Fatalf("tampered signed pack: %+v", got)
	}
}

func TestManifestRulesRefuseUnsafePacks(t *testing.T) {
	cases := map[string]string{
		"script outside": strings.Replace(manifest, "holds.ps1", "../evil.ps1", 1),
		"not ps1":        strings.Replace(manifest, "holds.ps1", "holds.exe", 1),
		"page":           strings.Replace(manifest, "page: reports", "page: settings", 1),
		"module":         strings.Replace(manifest, "module: exo", "module: graph", 1),
		"destructive":    strings.Replace(manifest, "danger: read", "danger: destructive", 1),
		"name":           strings.Replace(manifest, "contoso-tools", "Contoso Tools", 1),
	}
	for name, m := range cases {
		root := t.TempDir()
		writePack(t, root, "p", m)
		if got := Load(root, Trust{})[0]; got.Status != Invalid {
			t.Errorf("%s: %+v", name, got)
		}
	}
}

// The sample pack in the repository must stay loadable.
func TestSamplePackIsValid(t *testing.T) {
	list := Load(filepath.Join("..", "..", "..", "packs"), Trust{})
	if len(list) != 1 || list[0].Status != Untrusted || len(list[0].Manifest.Actions) != 2 {
		t.Fatalf("sample pack: %+v", list)
	}
}

func TestLookAlikesKeyHijackEncodingAndDisable(t *testing.T) {
	// Two folders with one name: neither loads.
	root := t.TempDir()
	writePack(t, root, "a", manifest)
	writePack(t, root, "0-copy", manifest)
	for _, p := range Load(root, Trust{}) {
		if p.Status != Invalid {
			t.Fatalf("duplicate name must not load: %+v", p)
		}
	}

	// A user key that claims the project key's id is ignored.
	builtin := trustedKeys(nil)
	var id [8]byte
	for k := range builtin {
		id = k
	}
	pub, _, _ := ed25519.GenerateKey(rand.Reader)
	fake := base64.StdEncoding.EncodeToString(append(append([]byte("Ed"), id[:]...), pub...))
	if keys := trustedKeys([]string{fake}); len(keys[id]) != 1 || keys[id][0].name != "SwissKnife project" {
		t.Fatalf("project key shadowed: %+v", keys[id])
	}

	// Scripts must be UTF-8.
	root = t.TempDir()
	d := writePack(t, root, "p", manifest)
	_ = os.WriteFile(filepath.Join(d, "holds.ps1"), []byte{0xFF, 0xFE, 'G', 0}, 0o644)
	if p := Load(root, Trust{})[0]; p.Status != Invalid || !strings.Contains(p.Error, "UTF-8") {
		t.Fatalf("UTF-16 script: %+v", p)
	}

	// A disabled pack stays off even when signed.
	root = t.TempDir()
	d = writePack(t, root, "p", manifest)
	key := sign(t, d)
	if p := Load(root, Trust{Keys: []string{key}, Disabled: []string{"contoso-tools"}})[0]; p.Status != Disabled || p.Usable() {
		t.Fatalf("disabled pack: %+v", p)
	}
}
