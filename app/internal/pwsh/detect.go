package pwsh

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"
)

// Module names the app can use and install. Install only ever passes one of
// these to Install-Module.
const (
	ModuleExchange = "ExchangeOnlineManagement"
	ModuleTeams    = "MicrosoftTeams"
)

// Minimum versions: pwsh 7.2 is the oldest the modules still load on; the
// Exchange module needs 3.1 for -AccessToken (3.8 for Security & Compliance),
// Teams 5.0 for -AccessTokens.
var (
	minPwsh    = []int{7, 2}
	minModules = map[string][]int{ModuleExchange: {3, 1}, ModuleTeams: {5, 0}}
)

// Environment is what is installed on this machine.
type Environment struct {
	Exe     string            `json:"exe"`     // "" when PowerShell 7 is not found
	Version string            `json:"version"` // pwsh version
	Modules map[string]string `json:"modules"` // name → newest installed version
}

// Supports reports whether the module is installed in a usable version.
func (e Environment) Supports(module string) bool {
	v, ok := e.Modules[module]
	return ok && atLeast(v, minModules[module])
}

// PwshOK reports a usable PowerShell 7.
func (e Environment) PwshOK() bool { return e.Exe != "" && atLeast(e.Version, minPwsh) }

func atLeast(version string, min []int) bool {
	parts := strings.Split(strings.SplitN(version, "-", 2)[0], ".")
	for i, m := range min {
		if i >= len(parts) {
			return false
		}
		n, err := strconv.Atoi(parts[i])
		if err != nil {
			return false
		}
		if n != m {
			return n > m
		}
	}
	return true
}

// findExe locates pwsh: PATH first, then the default install folders.
func findExe() string {
	if p, err := exec.LookPath("pwsh"); err == nil {
		return p
	}
	var candidates []string
	switch runtime.GOOS {
	case "windows":
		for _, base := range []string{os.Getenv("ProgramFiles"), os.Getenv("ProgramW6432")} {
			if base != "" {
				candidates = append(candidates, filepath.Join(base, "PowerShell", "7", "pwsh.exe"))
			}
		}
	case "darwin":
		candidates = []string{"/usr/local/bin/pwsh", "/opt/homebrew/bin/pwsh", "/usr/local/microsoft/powershell/7/pwsh"}
	default:
		candidates = []string{"/usr/bin/pwsh", "/opt/microsoft/powershell/7/pwsh", "/snap/bin/pwsh"}
	}
	for _, c := range candidates {
		if st, err := os.Stat(c); err == nil && !st.IsDir() {
			return c
		}
	}
	return ""
}

const detectScript = `$m = @{}; Get-Module -ListAvailable -Name ExchangeOnlineManagement,MicrosoftTeams | ForEach-Object { $v = $_.Version.ToString(); if (-not $m[$_.Name] -or [version]$v -gt [version]$m[$_.Name]) { $m[$_.Name] = $v } }; @{ version = $PSVersionTable.PSVersion.ToString(); modules = $m } | ConvertTo-Json -Compress`

// Detector caches the environment; detection spawns pwsh, so it is not free.
type Detector struct {
	mu   sync.Mutex
	env  *Environment
	at   time.Time
	find func() string
	run  func(ctx context.Context, exe string, args ...string) ([]byte, error)
}

func NewDetector() *Detector {
	return &Detector{find: findExe, run: runPwsh}
}

func runPwsh(ctx context.Context, exe string, args ...string) ([]byte, error) {
	cmd := exec.CommandContext(ctx, exe, append([]string{"-NoLogo", "-NoProfile", "-NonInteractive"}, args...)...)
	hideWindow(cmd)
	return cmd.Output()
}

// Get returns the cached environment, detecting at most every 10 minutes.
func (d *Detector) Get(ctx context.Context) Environment {
	d.mu.Lock()
	defer d.mu.Unlock()
	if d.env != nil && time.Since(d.at) < 10*time.Minute {
		return *d.env
	}
	env := Environment{Exe: d.find(), Modules: map[string]string{}}
	if env.Exe != "" {
		cctx, cancel := context.WithTimeout(ctx, 30*time.Second)
		defer cancel()
		out, err := d.run(cctx, env.Exe, "-Command", detectScript)
		if err != nil {
			// A timeout or a cancelled check says nothing about what is
			// installed: report it, but do not remember it.
			return env
		}
		var r struct {
			Version string            `json:"version"`
			Modules map[string]string `json:"modules"`
		}
		if json.Unmarshal(out, &r) == nil {
			env.Version = r.Version
			if r.Modules != nil {
				env.Modules = r.Modules
			}
		}
	}
	d.env, d.at = &env, time.Now()
	return env
}

// Invalidate forces the next Get to detect again (after an install).
func (d *Detector) Invalidate() {
	d.mu.Lock()
	d.env = nil
	d.mu.Unlock()
}

// Install installs a known module for the current user.
func (d *Detector) Install(ctx context.Context, module string) error {
	if _, ok := minModules[module]; !ok {
		return fmt.Errorf("unknown module %q", module)
	}
	env := d.Get(ctx)
	if !env.PwshOK() {
		return fmt.Errorf("PowerShell 7.2 or later is required")
	}
	// The name is one of two constants, so composing the command is safe.
	// PSGallery explicitly: another registered repository must not win with
	// a higher version number. -Force answers the untrusted-repository prompt
	// for this install only, without changing the user's trust settings.
	script := "Install-Module -Name " + module + " -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop"
	cctx, cancel := context.WithTimeout(ctx, 10*time.Minute)
	defer cancel()
	_, err := d.run(cctx, env.Exe, "-Command", script)
	d.Invalidate()
	var ee *exec.ExitError
	if errors.As(err, &ee) && len(ee.Stderr) > 0 {
		return fmt.Errorf("install %s: %s", module, strings.TrimSpace(string(ee.Stderr)))
	}
	if err != nil {
		return fmt.Errorf("install %s: %w", module, err)
	}
	return nil
}
