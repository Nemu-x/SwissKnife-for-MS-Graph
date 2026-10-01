//go:build !windows

package services

import "errors"

// errUpdateWindowsOnly is returned by the platform hooks on non-Windows builds;
// the in-app installer flow only exists for the Windows NSIS package.
var errUpdateWindowsOnly = errors.New("in-app update is available on Windows only — use the releases page")

func launchElevated(string) error { return errUpdateWindowsOnly }

func revealInFolder(string) error { return errUpdateWindowsOnly }
