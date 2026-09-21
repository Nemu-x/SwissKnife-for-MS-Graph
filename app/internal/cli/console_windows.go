//go:build windows

package cli

import (
	"os"

	"golang.org/x/sys/windows"
)

const attachParentProcess = ^uint32(0) // ATTACH_PARENT_PROCESS = (DWORD)-1

var (
	kernel32          = windows.NewLazySystemDLL("kernel32.dll")
	procAttachConsole = kernel32.NewProc("AttachConsole")
)

// attachConsole connects the process to the console it was launched from.
// The GUI build is linked with -H windowsgui, so a plain start has no console
// and no usable std handles; attaching to the parent process (cmd.exe,
// PowerShell) lets output land in the caller's terminal. Handles that are
// already valid (stdout redirected to a file or a pipe) are left untouched,
// otherwise piping `--json` into ConvertFrom-Json would break. Any failure is
// silent: there is nothing to print to yet.
func attachConsole() {
	if err := procAttachConsole.Find(); err != nil {
		return
	}
	if r, _, _ := procAttachConsole.Call(uintptr(attachParentProcess)); r == 0 {
		return
	}
	if !usable(windows.STD_INPUT_HANDLE) {
		reopen("CONIN$", &os.Stdin)
	}
	if !usable(windows.STD_OUTPUT_HANDLE) {
		reopen("CONOUT$", &os.Stdout)
	}
	if !usable(windows.STD_ERROR_HANDLE) {
		reopen("CONOUT$", &os.Stderr)
	}
}

// usable reports whether the standard handle already points at something
// (a redirected file or pipe) that must be preserved.
func usable(std uint32) bool {
	h, err := windows.GetStdHandle(std)
	if err != nil || h == 0 || h == windows.InvalidHandle {
		return false
	}
	_, err = windows.GetFileType(h)
	return err == nil
}

func reopen(name string, target **os.File) {
	h, err := windows.Open(name, windows.O_RDWR, 0)
	if err != nil {
		return
	}
	*target = os.NewFile(uintptr(h), name)
}
