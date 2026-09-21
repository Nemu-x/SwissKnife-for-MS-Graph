//go:build windows

package services

import (
	"errors"
	"os/exec"

	"golang.org/x/sys/windows"
)

// launchElevated starts the installer through the shell with the "runas" verb
// so Windows raises a UAC prompt. The per-machine NSIS installer needs
// elevation and os/exec cannot ask for it (ERROR_ELEVATION_REQUIRED, 740).
// A declined prompt (ERROR_CANCELLED, 1223) maps to ErrUpdateDeclined.
func launchElevated(exe string) error {
	verb, err := windows.UTF16PtrFromString("runas")
	if err != nil {
		return err
	}
	file, err := windows.UTF16PtrFromString(exe)
	if err != nil {
		return err
	}
	args, err := windows.UTF16PtrFromString("/S")
	if err != nil {
		return err
	}
	err = windows.ShellExecute(0, verb, file, args, nil, windows.SW_HIDE)
	if err == nil {
		return nil
	}
	if errors.Is(err, windows.ERROR_CANCELLED) {
		return ErrUpdateDeclined
	}
	return err
}

// revealInFolder opens Explorer with the given file selected.
func revealInFolder(path string) error {
	return exec.Command("explorer.exe", "/select,"+path).Start()
}
