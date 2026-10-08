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
	path, _, err := fetch(ctx, func(v string) string { return "powershell-" + v + "-osx-" + arch + ".pkg" }, progress)
	if err != nil {
		return err
	}
	defer func() { _ = os.RemoveAll(filepath.Dir(path)) }()
	progress("install", -1)
	// installer(8) checks the package's signature itself.
	q := strings.NewReplacer(`\`, `\\`, `"`, `\"`).Replace(path)
	script := `do shell script "/usr/sbin/installer -pkg " & quoted form of "` + q + `" & " -target /" with administrator privileges`
	out, err := exec.CommandContext(ctx, "/usr/bin/osascript", "-e", script).CombinedOutput()
	if err != nil {
		if strings.Contains(string(out), "-128") || strings.Contains(strings.ToLower(string(out)), "cancel") {
			return ErrDeclined
		}
		return fmt.Errorf("installer: %s", strings.TrimSpace(string(out)))
	}
	return nil
}
