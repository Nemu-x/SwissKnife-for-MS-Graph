package services

import (
	"crypto/ed25519"
	"crypto/rand"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	goruntime "runtime"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/jedisct1/go-minisign"
)

// testSigner is a throwaway minisign key pair.
type testSigner struct {
	sk  minisign.PrivateKey
	pub string // base64 public key, as in minisign.pub
}

func newTestSigner(t *testing.T, id byte) *testSigner {
	t.Helper()
	pub, priv, err := ed25519.GenerateKey(rand.Reader)
	if err != nil {
		t.Fatal(err)
	}
	s := &testSigner{}
	s.sk.SignatureAlgorithm = [2]byte{'E', 'd'}
	s.sk.KeyId = [8]byte{id, 2, 3, 4, 5, 6, 7, 8}
	copy(s.sk.SecretKey[:], priv)
	s.pub = base64.StdEncoding.EncodeToString(append(append([]byte("Ed"), s.sk.KeyId[:]...), pub...))
	return s
}

func (s *testSigner) sign(t *testing.T, data []byte, prehashed bool) []byte {
	t.Helper()
	sig, err := s.sk.Sign(data, minisign.SignOptions{Hashed: prehashed, TrustedComment: "release test"})
	if err != nil {
		t.Fatal(err)
	}
	return sig.Encode()
}

// fakeRelease is a GitHub "latest release" endpoint plus its assets.
type fakeRelease struct {
	srv            *httptest.Server
	installer      []byte
	installerName  string
	sums, sig      []byte
	noSig          bool
	installerFetch atomic.Int32
}

func newFakeRelease(t *testing.T) *fakeRelease {
	t.Helper()
	fr := &fakeRelease{
		installer:     []byte("MZ fake installer payload"),
		installerName: fmt.Sprintf("SwissKnifeGraph-t%d%s", time.Now().UnixNano(), installerAssetName()),
	}
	t.Cleanup(func() { _ = os.Remove(filepath.Join(os.TempDir(), fr.installerName)) })
	fr.srv = httptest.NewTLSServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		base := fr.srv.URL + "/dl/"
		switch r.URL.Path {
		case "/api":
			assets := []map[string]any{
				// The other architecture must never be picked.
				{"name": "SwissKnifeGraph-windows-other-installer.exe", "browser_download_url": base + "other", "size": 1},
				{"name": fr.installerName, "browser_download_url": base + "installer", "size": len(fr.installer)},
				{"name": checksumsAsset, "browser_download_url": base + "sums", "size": len(fr.sums)},
			}
			if !fr.noSig {
				assets = append(assets, map[string]any{"name": signatureAsset, "browser_download_url": base + "sig", "size": len(fr.sig)})
			}
			_ = json.NewEncoder(w).Encode(map[string]any{"tag_name": "v99.0.0", "assets": assets})
		case "/dl/installer":
			fr.installerFetch.Add(1)
			_, _ = w.Write(fr.installer)
		case "/dl/sums":
			_, _ = w.Write(fr.sums)
		case "/dl/sig":
			_, _ = w.Write(fr.sig)
		default:
			http.NotFound(w, r)
		}
	}))
	t.Cleanup(fr.srv.Close)

	// Point the updater at the fake server for this test only.
	oldAPI, oldClient, oldKey := releasesAPI, updateHTTPClient, releasePublicKey
	allowedAssetHosts["127.0.0.1"] = true
	releasesAPI = fr.srv.URL + "/api"
	c := fr.srv.Client()
	c.CheckRedirect = checkAssetRedirect
	updateHTTPClient = c
	t.Cleanup(func() {
		releasesAPI, updateHTTPClient, releasePublicKey = oldAPI, oldClient, oldKey
		delete(allowedAssetHosts, "127.0.0.1")
	})
	return fr
}

// signWith publishes SHA256SUMS.txt for the installer, signed by s.
func (fr *fakeRelease) signWith(t *testing.T, s *testSigner, prehashed bool) {
	sum := sha256.Sum256(fr.installer)
	fr.sums = []byte(fmt.Sprintf("%s  SwissKnifeGraph-darwin-universal.zip\n%s *%s\n", strings.Repeat("ab", 32), hex.EncodeToString(sum[:]), fr.installerName))
	fr.sig = s.sign(t, fr.sums, prehashed)
}

func checkAndDownload(t *testing.T, u *UpdateService, fr *fakeRelease) (string, error) {
	t.Helper()
	if _, err := u.Check(); err != nil {
		t.Fatal(err)
	}
	u.mu.Lock()
	rel := u.release
	u.mu.Unlock()
	if rel.installerName != fr.installerName {
		t.Fatalf("picked %q, want %q", rel.installerName, fr.installerName)
	}
	return u.Download(rel.installerURL, rel.installerName, int64(len(fr.installer)))
}

func TestUpdateDownloadVerifiesSignedChecksums(t *testing.T) {
	for _, prehashed := range []bool{false, true} {
		t.Run(fmt.Sprintf("prehashed=%v", prehashed), func(t *testing.T) {
			fr := newFakeRelease(t)
			key := newTestSigner(t, 1)
			releasePublicKey = key.pub
			fr.signWith(t, key, prehashed)

			u := NewUpdateService("v1.0.0")
			path, err := checkAndDownload(t, u, fr)
			if err != nil {
				t.Fatalf("verified download failed: %v", err)
			}
			if got, _ := os.ReadFile(path); string(got) != string(fr.installer) {
				t.Fatal("installer content differs")
			}
			if u.verified[filepath.Clean(path)] == "" {
				t.Fatal("download not recorded as verified")
			}
		})
	}
}

func TestUpdateDownloadRefusesUnverifiedInstallers(t *testing.T) {
	cases := []struct {
		name  string
		setup func(t *testing.T, fr *fakeRelease, key *testSigner)
		want  string
	}{
		{"hash mismatch", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.signWith(t, key, true)
			fr.installer = []byte("MZ tampered installer!!!!")
		}, "does not match"},
		{"foreign key", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.signWith(t, newTestSigner(t, 1), true) // same key id, different key
		}, "signature is invalid"},
		{"tampered sums", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.signWith(t, key, false)
			fr.sums = append(fr.sums, '\n')
		}, "signature is invalid"},
		{"garbage signature", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.signWith(t, key, false)
			fr.sig = []byte("not a signature")
		}, "malformed"},
		{"no signature asset", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.signWith(t, key, false)
			fr.noSig = true
		}, "no signed checksums"},
		{"installer not listed", func(t *testing.T, fr *fakeRelease, key *testSigner) {
			fr.sums = []byte(strings.Repeat("cd", 32) + "  other.exe\n")
			fr.sig = key.sign(t, fr.sums, true)
		}, "no entry"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			fr := newFakeRelease(t)
			key := newTestSigner(t, 1)
			releasePublicKey = key.pub
			tc.setup(t, fr, key)

			u := NewUpdateService("v1.0.0")
			_, err := checkAndDownload(t, u, fr)
			if err == nil || !strings.Contains(err.Error(), tc.want) {
				t.Fatalf("err = %v, want %q", err, tc.want)
			}
			if _, statErr := os.Stat(filepath.Join(os.TempDir(), fr.installerName)); statErr == nil {
				t.Fatal("an unverified installer was left on disk")
			}
			if tc.name != "hash mismatch" && fr.installerFetch.Load() != 0 {
				t.Fatal("the installer was downloaded although the checksums were not trusted")
			}
		})
	}
}

func TestUpdateDownloadOnlyTheCheckedAsset(t *testing.T) {
	fr := newFakeRelease(t)
	key := newTestSigner(t, 1)
	releasePublicKey = key.pub
	fr.signWith(t, key, true)
	u := NewUpdateService("v1.0.0")
	// No Check yet: nothing is trusted.
	if _, err := u.Download(fr.srv.URL+"/dl/installer", fr.installerName, 0); err == nil {
		t.Fatal("download without a check must be refused")
	}
	if _, err := u.Check(); err != nil {
		t.Fatal(err)
	}
	if _, err := u.Download(fr.srv.URL+"/dl/other", fr.installerName, 0); err == nil {
		t.Fatal("an asset other than the checked installer must be refused")
	}
	if _, err := u.Download("http://github.com/x", fr.installerName, 0); err == nil {
		t.Fatal("plain http must be refused")
	}
}

// Apply must re-hash right before launching: a file changed after the
// verified download, or never verified at all, is not run.
func TestApplyRefusesChangedOrUnverifiedInstaller(t *testing.T) {
	fr := newFakeRelease(t)
	key := newTestSigner(t, 1)
	releasePublicKey = key.pub
	fr.signWith(t, key, true)
	u := NewUpdateService("v1.0.0")
	path, err := checkAndDownload(t, u, fr)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("MZ swapped after verification"), 0o644); err != nil {
		t.Fatal(err)
	}
	err = u.Apply(path)
	if err == nil {
		t.Fatal("a changed installer must not be launched")
	}
	if goruntime.GOOS == "windows" && !strings.Contains(err.Error(), "changed after it was verified") {
		t.Fatalf("err = %v", err)
	}

	u2 := NewUpdateService("v1.0.0")
	err = u2.Apply(path)
	if err == nil {
		t.Fatal("an unverified installer must not be launched")
	}
	if goruntime.GOOS == "windows" && !strings.Contains(err.Error(), "not verified") {
		t.Fatalf("err = %v", err)
	}
}

func TestSafeReleasesURL(t *testing.T) {
	cases := map[string]string{
		"": releasesPage,
		"https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/tag/v2.0.0":      "https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/tag/v2.0.0",
		"https://github.com/Nemu-x/SwissKnife-for-MS-Graph/wiki/Teams-Notifications": "https://github.com/Nemu-x/SwissKnife-for-MS-Graph/wiki/Teams-Notifications",
		"http://github.com/Nemu-x/x":           releasesPage,
		"https://github.com/evil/x":            releasesPage,
		"https://github.com.evil.io/Nemu-x/x":  releasesPage,
		"https://user@github.com/Nemu-x/x":     releasesPage,
		"file:///C:/Windows/System32/calc.exe": releasesPage,
		"javascript:alert(1)":                  releasesPage,
		`C:\Windows\System32\cmd.exe`:          releasesPage,
	}
	for in, want := range cases {
		if got := safeReleasesURL(in); got != want {
			t.Errorf("safeReleasesURL(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestCleanupStaleInstallers(t *testing.T) {
	dir := t.TempDir()
	old := filepath.Join(dir, "SwissKnifeGraph-windows-amd64-installer.exe")
	fresh := filepath.Join(dir, "SwissKnifeGraph-windows-arm64-installer.exe")
	other := filepath.Join(dir, "Other-windows-amd64-installer.exe")
	for _, p := range []string{old, fresh, other} {
		if err := os.WriteFile(p, []byte("MZ"), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	past := time.Now().Add(-48 * time.Hour)
	for _, p := range []string{old, other} {
		if err := os.Chtimes(p, past, past); err != nil {
			t.Fatal(err)
		}
	}
	cleanupStaleInstallers(dir, 24*time.Hour)
	if _, err := os.Stat(old); !os.IsNotExist(err) {
		t.Error("stale installer kept")
	}
	for _, p := range []string{fresh, other} {
		if _, err := os.Stat(p); err != nil {
			t.Errorf("%s removed: %v", filepath.Base(p), err)
		}
	}
}
