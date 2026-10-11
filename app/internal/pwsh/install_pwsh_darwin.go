package pwsh

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
)

// Method: the signed .pkg, installed with the system password prompt.
func Method() InstallMethod { return InstallMethod{Auto: true, How: "pkg", URL: DocsURL} }

func installPlatform(ctx context.Context, progress InstallProgress) error {
	arch := "x64"
	if runtime.GOARCH == "arm64" {
		arch = "arm64"
	}
	path, sum, err := fetch(ctx, func(v string) string { return "powershell-" + v + "-osx-" + arch + ".pkg" }, progress)
	if err != nil {
		return err
	}
	defer func() { _ = os.RemoveAll(filepath.Dir(path)) }()
	// installer(8) running as root does not insist on a valid signature, so
	// check it here: the package must be signed by Microsoft's Developer ID.
	out, err := exec.CommandContext(ctx, "/usr/sbin/pkgutil", "--check-signature", path).CombinedOutput()
	if err != nil {
		return fmt.Errorf("the installer's signature is not valid — not installing it: %s", strings.TrimSpace(string(out)))
	}
	if err := checkPkgSignatureOutput(string(out)); err != nil {
		return fmt.Errorf("the installer's signature is not valid — not installing it: %w", err)
	}
	// The file sits in a folder this user can write: check it once more right
	// before the elevated install (macOS has no share-mode lock to hold).
	if err := hashFileMatches(path, sum); err != nil {
		return err
	}
	progress("install", -1)
	q :=strings.NewReplacer(`\`, `\\`, `"`, `\"`).Replace(path)
	script := `do shell script "/usr/sbin/installer -pkg " & quoted form of "` + q + `" & " -target /" with administrator privileges`
	out, err = exec.CommandContext(ctx, "/usr/bin/osascript", "-e", script).CombinedOutput()
	if err != nil {
		if strings.Contains(string(out), "-128") || strings.Contains(strings.ToLower(string(out)), "cancel") {
			return ErrDeclined
		}
		return fmt.Errorf("installer: %s", strings.TrimSpace(string(out)))
	}
	return nil
}
