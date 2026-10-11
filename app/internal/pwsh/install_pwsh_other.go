//go:build !windows && !darwin

package pwsh

import (
	"context"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"strings"
)

// Method: snap through the desktop's admin prompt (pkexec) when both exist,
// otherwise the commands for this distribution.
func Method() InstallMethod {
	if hasSnapAndPkexec() {
		return InstallMethod{Auto: true, How: "snap", URL: DocsURL}
	}
	return InstallMethod{How: "manual", Commands: distroCommands(), URL: DocsURL}
}

func hasSnapAndPkexec() bool {
	_, e1 := exec.LookPath("snap")
	_, e2 := exec.LookPath("pkexec")
	return e1 == nil && e2 == nil
}

func distroCommands() []string {
	id := ""
	if b, err := os.ReadFile("/etc/os-release"); err == nil {
		for _, l := range strings.Split(string(b), "\n") {
			if strings.HasPrefix(l, "ID=") || strings.HasPrefix(l, "ID_LIKE=") {
				id += " " + strings.Trim(strings.SplitN(l, "=", 2)[1], `"`)
			}
		}
	}
	switch {
	case strings.Contains(id, "arch"):
		return []string{"yay -S powershell-bin"}
	case strings.Contains(id, "fedora"), strings.Contains(id, "rhel"):
		return []string{
			"sudo dnf install -y https://packages.microsoft.com/config/rhel/9/packages-microsoft-prod.rpm",
			"sudo dnf install -y powershell",
		}
	default:
		return []string{"sudo snap install powershell --classic"}
	}
}

func installPlatform(ctx context.Context, progress InstallProgress) error {
	if !hasSnapAndPkexec() {
		return errors.New("install PowerShell with the commands shown")
	}
	progress("install", -1)
	out, err := exec.CommandContext(ctx, "pkexec", "snap", "install", "powershell", "--classic").CombinedOutput()
	if err != nil {
		var ee *exec.ExitError
		if errors.As(err, &ee) && (ee.ExitCode() == 126 || ee.ExitCode() == 127) {
			return ErrDeclined
		}
		if msg := strings.TrimSpace(string(out)); msg != "" {
			return fmt.Errorf("snap: %s", msg)
		}
		return fmt.Errorf("snap: %w", err)
	}
	return nil
}
