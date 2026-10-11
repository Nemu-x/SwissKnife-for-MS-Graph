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
// are signed with: the full identity, Team ID included (a name prefix alone
// could match another developer's certificate).
const microsoftInstallerCert = "Developer ID Installer: Microsoft Corporation (UBF8T346G9)"

// elevatedPkgInstall is the shell script run with administrator rights by
// install_pwsh_darwin.go, with the downloaded package as $1. The user's own
// copy can be swapped while the password prompt is up, so the script copies
// it into a fresh root-owned folder, then checks that copy's SHA-256 and
// signature before installing exactly that copy.
func elevatedPkgInstall(wantSHA256 string) string {
	return `set -e
d=$(/usr/bin/mktemp -d /private/var/tmp/skg-pwsh.XXXXXX)
trap '/bin/rm -rf "$d"' EXIT
/bin/cp "$1" "$d/pwsh.pkg"
sum=$(/usr/bin/shasum -a 256 "$d/pwsh.pkg" | /usr/bin/cut -d ' ' -f 1)
if [ "$sum" != "` + wantSHA256 + `" ]; then echo "the installer changed after it was downloaded" >&2; exit 3; fi
if ! /usr/sbin/pkgutil --check-signature "$d/pwsh.pkg" | /usr/bin/grep -qF "1. ` + microsoftInstallerCert + `"; then echo "the installer is not signed by Microsoft" >&2; exit 4; fi
/usr/sbin/installer -pkg "$d/pwsh.pkg" -target /
`
}

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
			if strings.TrimPrefix(line, "1. ") == microsoftInstallerCert {
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
