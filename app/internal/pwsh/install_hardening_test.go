package pwsh

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

const pkgutilMicrosoft = `Package "powershell-7.6.6-osx-arm64.pkg":
   Status: signed by a developer certificate issued by Apple for distribution
   Notarization: trusted by the Apple notary service
   Signed with a trusted timestamp on: 2026-05-01 10:00:00 +0000
   Certificate Chain:
    1. Developer ID Installer: Microsoft Corporation (UBF8T346G9)
       Expires: 2030-01-01 00:00:00 +0000
    2. Developer ID Certification Authority
    3. Apple Root CA
`

func TestCheckPkgSignatureOutput(t *testing.T) {
	if err := checkPkgSignatureOutput(pkgutilMicrosoft); err != nil {
		t.Fatalf("Microsoft-signed package refused: %v", err)
	}
	bad := map[string]string{
		"other signer":  strings.Replace(pkgutilMicrosoft, "Microsoft Corporation", "Evil Corp", 1),
		"unsigned":      "Package \"x.pkg\":\n   Status: no signature\n",
		"self-signed":   strings.Replace(pkgutilMicrosoft, "signed by a developer certificate issued by Apple for distribution", "signed by untrusted certificate", 1),
		"no chain":      "   Status: signed by a developer certificate issued by Apple for distribution\n",
		"name not leaf": strings.Replace(strings.Replace(pkgutilMicrosoft, "1. Developer ID Installer: Microsoft Corporation", "1. Developer ID Installer: Evil", 1), "2. Developer ID Certification Authority", "2. Developer ID Installer: Microsoft Corporation", 1),
	}
	for name, out := range bad {
		if err := checkPkgSignatureOutput(out); err == nil {
			t.Errorf("%s: accepted", name)
		}
	}
}

func TestHashFileMatches(t *testing.T) {
	p := filepath.Join(t.TempDir(), "x.pkg")
	_ = os.WriteFile(p, []byte("pkg"), 0o600)
	sum := sha256.Sum256([]byte("pkg"))
	if err := hashFileMatches(p, hex.EncodeToString(sum[:])); err != nil {
		t.Fatal(err)
	}
	_ = os.WriteFile(p, []byte("swapped"), 0o600)
	if err := hashFileMatches(p, hex.EncodeToString(sum[:])); err == nil {
		t.Fatal("a changed file must be refused")
	}
}

func TestStableTag(t *testing.T) {
	for tag, want := range map[string]bool{
		"v7.6.6": true, "v8.0.0": true, "v10.1.2": true,
		"v6.2.7": false, "v7.6.0-preview.1": false, "7.6.6": false, "v7.6": false,
		"v7.6.6/../../evil": false, "": false,
	} {
		if got := stableTag(tag); got != want {
			t.Errorf("stableTag(%q) = %v, want %v", tag, got, want)
		}
	}
}

func TestVersionedExesPrefersHighestMajor(t *testing.T) {
	root := t.TempDir()
	for _, v := range []string{"7", "8", "7-preview", "notes"} {
		_ = os.MkdirAll(filepath.Join(root, v), 0o755)
		_ = os.WriteFile(filepath.Join(root, v, "pwsh.exe"), []byte("x"), 0o755)
	}
	got := versionedExes(root, "pwsh.exe")
	var dirs []string
	for _, p := range got {
		dirs = append(dirs, filepath.Base(filepath.Dir(p)))
	}
	if strings.Join(dirs, ",") != "8,7,7-preview,notes" {
		t.Fatalf("order %v", dirs)
	}
	if versionedExes(filepath.Join(root, "missing"), "pwsh.exe") != nil {
		t.Fatal("missing root must yield nothing")
	}
}

// A download that stops sending bytes ends with an error instead of hanging.
func TestFetchStallTimesOut(t *testing.T) {
	pkg := []byte("pretend this is an msi")
	sum := sha256.Sum256(pkg)
	release := make(chan struct{})
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/info":
			_, _ = w.Write([]byte(`{"ReleaseTag":"v8.0.1"}`))
		case strings.HasSuffix(r.URL.Path, "/hashes.sha256"):
			_, _ = w.Write([]byte(hex.EncodeToString(sum[:]) + " *PowerShell-8.0.1-win-x64.msi\n"))
		default:
			w.Header().Set("Content-Length", "1000")
			_, _ = w.Write(pkg[:4])
			w.(http.Flusher).Flush()
			select { // stall until the client gives up or the test ends
			case <-r.Context().Done():
			case <-release:
			}
		}
	}))
	t.Cleanup(func() { close(release); srv.Close() })
	oldInfo, oldBase, oldIdle := releaseInfoURL, downloadBase, packageIdleTimeout
	releaseInfoURL, downloadBase, packageIdleTimeout = srv.URL+"/info", srv.URL+"/", 200*time.Millisecond
	t.Cleanup(func() { releaseInfoURL, downloadBase, packageIdleTimeout = oldInfo, oldBase, oldIdle })

	done := make(chan error, 1)
	go func() {
		_, _, err := fetch(context.Background(), func(v string) string { return "PowerShell-" + v + "-win-x64.msi" }, nil)
		done <- err
	}()
	select {
	case err := <-done:
		if err == nil || !strings.Contains(err.Error(), "stalled") {
			t.Fatalf("err = %v, want a stall error", err)
		}
	case <-time.After(10 * time.Second):
		t.Fatal("a stalled download must not hang the install")
	}
}
