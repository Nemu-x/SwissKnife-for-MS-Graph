//go:build !windows

package pwsh

import "os/exec"

func hideWindow(*exec.Cmd) {}
