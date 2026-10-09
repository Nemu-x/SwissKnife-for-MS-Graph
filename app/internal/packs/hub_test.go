package packs

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func copyDir(t *testing.T, src, dst string) {
	t.Helper()
	_ = filepath.WalkDir(src, func(p string, d os.DirEntry, err error) error {
		rel, _ := filepath.Rel(src, p)
		if d.IsDir() {
			return os.MkdirAll(filepath.Join(dst, rel), 0o755)
		}
		b, _ := os.ReadFile(p)
		return os.WriteFile(filepath.Join(dst, rel), b, 0o644)
	})
}

func TestHubIndexInstallAndTamper(t *testing.T) {
	hubDir := t.TempDir()
	copyDir(t, filepath.Join("..", "..", "..", "packs"), filepath.Join(hubDir, "packs"))
	_ = os.Remove(filepath.Join(hubDir, "packs", "README.md"))
	idx, err := BuildIndex(hubDir)
	if err != nil || len(idx.Packs) < 2 {
		t.Fatalf("index %+v %v", idx, err)
	}
	b, _ := json.Marshal(idx)
	_ = os.WriteFile(filepath.Join(hubDir, "index.json"), b, 0o644)

	srv := httptest.NewTLSServer(http.FileServer(http.Dir(hubDir)))
	t.Cleanup(srv.Close)
	old := hubClient
	hubClient = srv.Client()
	t.Cleanup(func() { hubClient = old })

	got, err := FetchIndex(context.Background(), srv.URL)
	if err != nil {
		t.Fatal(err)
	}
	var entry HubEntry
	for _, e := range got.Packs {
		if e.Name == "security-basics" {
			entry = e
		}
	}
	if entry.Kind != "workflow" || entry.Category != "other" && entry.Category == "" {
		t.Fatalf("entry %+v", entry)
	}
	root := t.TempDir()
	dir, err := Install(context.Background(), srv.URL, root, entry)
	if err != nil {
		t.Fatal(err)
	}
	if p := Load(root, Trust{})[0]; p.Status != Untrusted || p.Digest != entry.Digest || p.Dir != dir {
		t.Fatalf("installed %+v", p)
	}

	// A file changed on the hub after the index was built is refused.
	_ = os.WriteFile(filepath.Join(hubDir, "packs", "security-basics", "manifest.yaml"), []byte("name: security-basics\n"), 0o644)
	if _, err := Install(context.Background(), srv.URL, root, entry); err == nil || !strings.Contains(err.Error(), "does not match") {
		t.Fatalf("tampered pack: %v", err)
	}
	// The earlier install is untouched.
	if p := Load(root, Trust{})[0]; p.Digest != entry.Digest {
		t.Fatal("a failed update must keep the installed pack")
	}

	// Paths cannot leave the pack folder.
	bad := entry
	bad.Files = []string{"../../etc/passwd"}
	if _, err := Install(context.Background(), srv.URL, root, bad); err == nil {
		t.Fatal("a path outside the pack must be refused")
	}
	if _, err := FetchIndex(context.Background(), "http://example.com/"); err == nil {
		t.Fatal("plain http must be refused")
	}
}

func TestHubSafetyRules(t *testing.T) {
	if !Newer("1.2.0", "1.1.9") || Newer("1.0.0", "1.0.0") || Newer("1.0.0", "1.10.0") || !Newer("2.0", "1.9.9") {
		t.Fatal("version order")
	}
	for _, bad := range []string{"con", "a/NUL.txt", "x.", "a/../b", ".hidden", "a\b"} {
		if safeRel(bad) {
			t.Errorf("%q must be refused", bad)
		}
	}
	// Staging and backup folders are never packs.
	root := t.TempDir()
	writePack(t, root, ".staging-x-1", manifest)
	writePack(t, root, ".trash-contoso-tools", manifest)
	writePack(t, root, "contoso", manifest)
	if l := Load(root, Trust{}); len(l) != 1 || l[0].Status != Untrusted {
		t.Fatalf("hidden folders must be skipped: %+v", l)
	}

	// The hub may not redirect to plain http.
	srv := httptest.NewTLSServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		http.Redirect(w, r, "http://example.com/index.json", http.StatusFound)
	}))
	t.Cleanup(srv.Close)
	old := hubClient
	c := srv.Client()
	c.CheckRedirect = hubClient.CheckRedirect
	hubClient = c
	t.Cleanup(func() { hubClient = old })
	if _, err := FetchIndex(context.Background(), srv.URL); err == nil || !strings.Contains(err.Error(), "redirect") {
		t.Fatalf("redirect to http: %v", err)
	}
}

// A hub entry whose files say another name, or a folder holding another
// pack, is refused.
func TestHubInstallRefusesNameMismatch(t *testing.T) {
	hubDir := t.TempDir()
	writePack(t, filepath.Join(hubDir, "packs"), "tools", manifest) // manifest says contoso-tools
	d, _ := Digest(filepath.Join(hubDir, "packs", "tools"))
	srv := httptest.NewTLSServer(http.FileServer(http.Dir(hubDir)))
	t.Cleanup(srv.Close)
	old := hubClient
	hubClient = srv.Client()
	t.Cleanup(func() { hubClient = old })
	e := HubEntry{Name: "tools", Path: "packs/tools", Files: []string{"holds.ps1", "manifest.yaml"}, Digest: d}
	if _, err := Install(context.Background(), srv.URL, t.TempDir(), e); err == nil || !strings.Contains(err.Error(), "named") {
		t.Fatalf("name mismatch: %v", err)
	}
}
