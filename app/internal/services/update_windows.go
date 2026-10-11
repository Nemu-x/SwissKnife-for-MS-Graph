//go:build windows

package services

import (
	"errors"
	"os"
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
	cmd := exec.Command("explorer.exe", "/select,"+path)
	if err := cmd.Start(); err != nil {
		return err
	}
	go func() { _ = cmd.Wait() }() // reap the child
	return nil
}

// openInstallerLocked opens the installer for reading with a share mode that
// denies writers and rename/delete while the handle is open: Apply hashes the
// file and launches it within this handle's lifetime, so nothing can swap the
// verified bytes in between. Readers (the installer itself) still work.
func openInstallerLocked(path string) (*os.File, error) {
	p, err := windows.UTF16PtrFromString(path)
	if err != nil {
		return nil, err
	}
	h, err := windows.CreateFile(p, windows.GENERIC_READ, windows.FILE_SHARE_READ, nil,
		windows.OPEN_EXISTING, windows.FILE_ATTRIBUTE_NORMAL, 0)
	if err != nil {
		return nil, err
	}
	return os.NewFile(uintptr(h), path), nil
}
