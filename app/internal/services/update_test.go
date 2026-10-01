package services

import (
	"os"
	"path/filepath"
	goruntime "runtime"
	"strings"
	"testing"
)

// TestApplyPathGuards: Apply and RevealInstaller must refuse anything that is
// not a release installer inside the OS temp directory before any launch. On
// non-Windows both refuse outright (the flow only exists for the NSIS package).
func TestApplyPathGuards(t *testing.T) {
	u := NewUpdateService("v1.0.0")

	// A real "-installer.exe" outside the temp dir: must never be launched.
	outside := filepath.Join(t.TempDir(), "sub", "SwissKnifeGraph-windows-amd64-installer.exe")
	if err := os.MkdirAll(filepath.Dir(outside), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(outside, []byte("MZ"), 0o644); err != nil {
		t.Fatal(err)
	}

	cases := []struct {
		name, path, want string
	}{
		{"not an installer", filepath.Join(os.TempDir(), "notes.txt"), "not a release installer"},
		{"outside temp dir", outside, "update download location"},
		{"missing file", filepath.Join(os.TempDir(), "definitely-missing-installer.exe"), ""},
	}
	for _, tc := range cases {
		for _, fn := range []struct {
			name string
			call func(string) error
		}{{"Apply", u.Apply}, {"RevealInstaller", u.RevealInstaller}} {
			err := fn.call(tc.path)
			if err == nil {
				t.Fatalf("%s(%s): expected error, got nil", fn.name, tc.name)
			}
			if goruntime.GOOS != "windows" {
				if !strings.Contains(err.Error(), "Windows only") {
					t.Errorf("%s(%s) on %s: want Windows-only error, got %q", fn.name, tc.name, goruntime.GOOS, err)
				}
				continue
			}
			if tc.want != "" && !strings.Contains(err.Error(), tc.want) {
				t.Errorf("%s(%s): want %q in error, got %q", fn.name, tc.name, tc.want, err)
			}
		}
	}
}

// TestErrUpdateDeclinedEnvelope: the declined-UAC sentinel must cross the Wails
// boundary as an OpError with a stable code the frontend keys on.
func TestErrUpdateDeclinedEnvelope(t *testing.T) {
	s := ErrUpdateDeclined.Error()
	if !strings.HasPrefix(s, "operr:") || !strings.Contains(s, `"code":"update_declined"`) {
		t.Fatalf("unexpected envelope: %s", s)
	}
}

func TestIsNewerVersion(t *testing.T) {
	cases := []struct {
		latest, current string
		want            bool
	}{
		{"v1.1.0", "v1.0.0", true},
		{"v1.0.0", "v1.0.0", false},
		{"v0.9.9", "v1.0.0", false},
		{"v1.0.1", "1.0.0-dev", true},
		{"nightly", "v1.0.0", false},
	}
	for _, tc := range cases {
		if got := isNewer(tc.latest, tc.current); got != tc.want {
			t.Errorf("isNewer(%q,%q)=%v want %v", tc.latest, tc.current, got, tc.want)
		}
	}
}
