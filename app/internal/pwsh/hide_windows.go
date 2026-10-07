//go:build windows

package pwsh

import (
	"os/exec"
	"syscall"
)

// hideWindow keeps pwsh from flashing a console window over the GUI.
func hideWindow(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x08000000} // CREATE_NO_WINDOW
}
