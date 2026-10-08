package pwsh

import (
	"context"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"
	"unicode/utf16"
)

func utf16File(s string) []byte {
	u := utf16.Encode([]rune(s))
	b := []byte{0xFF, 0xFE}
	for _, r := range u {
		b = binary.LittleEndian.AppendUint16(b, r)
	}
	return b
}

// The package is downloaded, checked against the published (UTF-16) hash
// list, and refused when it does not match.
func TestFetchVerifiesThePublishedChecksum(t *testing.T) {
	pkg := []byte("pretend this is an msi")
	sum := sha256.Sum256(pkg)
	serve := pkg
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/info":
			_, _ = w.Write([]byte(`{"ReleaseTag":"v7.6.6"}`))
		case strings.HasSuffix(r.URL.Path, "/hashes.sha256"):
			_, _ = w.Write(utf16File("deadbeef *other.rpm\r\n" + hex.EncodeToString(sum[:]) + " *PowerShell-7.6.6-win-x64.msi\r\n"))
		case strings.HasSuffix(r.URL.Path, "/v7.6.6/PowerShell-7.6.6-win-x64.msi"):
			_, _ = w.Write(serve)
		default:
			http.NotFound(w, r)
		}
	}))
	t.Cleanup(srv.Close)
	oldInfo, oldBase := releaseInfoURL, downloadBase
	releaseInfoURL, downloadBase = srv.URL+"/info", srv.URL+"/"
	t.Cleanup(func() { releaseInfoURL, downloadBase = oldInfo, oldBase })

	name := func(v string) string { return "PowerShell-" + v + "-win-x64.msi" }
	var stages []string
	path, _, err := fetch(context.Background(), name, func(s string, _ int) { stages = append(stages, s) })
	if err != nil {
		t.Fatal(err)
	}
	b, _ := os.ReadFile(path)
	if string(b) != string(pkg) || stages[len(stages)-1] != "verify" {
		t.Fatalf("file %q stages %v", b, stages)
	}

	serve = []byte("tampered")
	if _, _, err := fetch(context.Background(), name, nil); err == nil || !strings.Contains(err.Error(), "checksum") {
		t.Fatalf("a package that does not match must be refused: %v", err)
	}
}
