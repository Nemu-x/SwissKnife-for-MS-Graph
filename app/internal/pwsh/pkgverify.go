package pwsh

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"io"
	"os"
	"strings"
)

// The checks below guard the macOS .pkg install (install_pwsh_darwin.go);
// they live in a file without a build tag so they are tested on every OS.

// microsoftInstallerCert is the leaf certificate PowerShell's macOS packages
// are signed with.
const microsoftInstallerCert = "Developer ID Installer: Microsoft Corporation"

// checkPkgSignatureOutput reads `pkgutil --check-signature` output: the
// package must carry a signature Apple's chain trusts, and its leaf (the
// first certificate of the chain) must be Microsoft's Developer ID Installer.
func checkPkgSignatureOutput(out string) error {
	if !strings.Contains(out, "Status: signed by a developer certificate issued by Apple") {
		return errors.New("the package is not signed with a trusted Developer ID")
	}
	inChain := false
	for _, line := range strings.Split(out, "\n") {
		line = strings.TrimSpace(line)
		if strings.HasPrefix(line, "Certificate Chain:") {
			inChain = true
			continue
		}
		if inChain && strings.HasPrefix(line, "1. ") {
			if strings.HasPrefix(strings.TrimPrefix(line, "1. "), microsoftInstallerCert) {
				return nil
			}
			return errors.New("the package is not signed by Microsoft")
		}
	}
	return errors.New("the package signature has no certificate chain")
}

// hashFileMatches re-reads path and compares its SHA-256 with want.
func hashFileMatches(path, want string) error {
	f, err := os.Open(path)
	if err != nil {
		return err
	}
	defer f.Close()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return err
	}
	if hex.EncodeToString(h.Sum(nil)) != want {
		return errors.New("the installer changed after it was downloaded — not installing it")
	}
	return nil
}
