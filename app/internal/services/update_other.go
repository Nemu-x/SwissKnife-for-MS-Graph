//go:build !windows

package services

import (
	"errors"
	"os/exec"
	"path/filepath"
	"runtime"
)

// errUpdateWindowsOnly is returned by the platform hooks on non-Windows builds;
// the in-app installer flow only exists for the Windows NSIS package.
var errUpdateWindowsOnly = errors.New("in-app update is available on Windows only — use the releases page")

func launchElevated(string) error { return errUpdateWindowsOnly }

// revealInFolder shows the file in Finder, or opens its folder elsewhere.
func revealInFolder(path string) error {
	cmd := exec.Command("xdg-open", filepath.Dir(path))
	if runtime.GOOS == "darwin" {
		cmd = exec.Command("open", "-R", path)
	}
	if err := cmd.Start(); err != nil {
		return err
	}
	go func() { _ = cmd.Wait() }() // reap the child
	return nil
}
