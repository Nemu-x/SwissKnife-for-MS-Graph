package pwsh

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"unsafe"

	"golang.org/x/sys/windows"
)

// Method: the MSI, installed for all users with one admin prompt.
func Method() InstallMethod { return InstallMethod{Auto: true, How: "msi", URL: DocsURL} }

func msiArch() string {
	if runtime.GOARCH == "arm64" {
		return "arm64"
	}
	return "x64"
}

func installPlatform(ctx context.Context, progress InstallProgress) error {
	path, err := fetch(ctx, func(v string) string { return "PowerShell-" + v + "-win-" + msiArch() + ".msi" }, progress)
	if err != nil {
		return err
	}
	defer func() { _ = os.RemoveAll(filepath.Dir(path)) }()
	// Beyond the checksum: the package must carry a valid code signature.
	if err := verifySignature(path); err != nil {
		return fmt.Errorf("the installer's signature is not valid — not installing it: %w", err)
	}
	progress("install", -1)
	// Microsoft Update keeps it current afterwards; ADD_PATH puts pwsh on PATH.
	args := fmt.Sprintf(`/i "%s" /qb- /norestart ADD_PATH=1 USE_MU=1 ENABLE_MU=1`, path)
	code, err := runElevatedWait("msiexec.exe", args)
	if err != nil {
		return err
	}
	switch code {
	case 0, 3010: // 3010: done, a restart finishes it
		return nil
	case 1602:
		return ErrDeclined
	default:
		return fmt.Errorf("the PowerShell installer ended with code %d", code)
	}
}

// verifySignature asks Windows whether the file's Authenticode signature is
// valid and trusted.
func verifySignature(path string) error {
	p, err := windows.UTF16PtrFromString(path)
	if err != nil {
		return err
	}
	file := &windows.WinTrustFileInfo{Size: uint32(unsafe.Sizeof(windows.WinTrustFileInfo{})), FilePath: p}
	data := &windows.WinTrustData{
		Size:                            uint32(unsafe.Sizeof(windows.WinTrustData{})),
		UIChoice:                        windows.WTD_UI_NONE,
		RevocationChecks:                windows.WTD_REVOKE_NONE,
		UnionChoice:                     windows.WTD_CHOICE_FILE,
		StateAction:                     windows.WTD_STATEACTION_VERIFY,
		FileOrCatalogOrBlobOrSgnrOrCert: unsafe.Pointer(file),
	}
	err = windows.WinVerifyTrustEx(windows.InvalidHWND, &windows.WINTRUST_ACTION_GENERIC_VERIFY_V2, data)
	data.StateAction = windows.WTD_STATEACTION_CLOSE
	_ = windows.WinVerifyTrustEx(windows.InvalidHWND, &windows.WINTRUST_ACTION_GENERIC_VERIFY_V2, data)
	return err
}

// shellExecuteInfo is SHELLEXECUTEINFOW (x/sys has no ShellExecuteEx).
type shellExecuteInfo struct {
	cbSize       uint32
	fMask        uint32
	hwnd         windows.Handle
	lpVerb       *uint16
	lpFile       *uint16
	lpParameters *uint16
	lpDirectory  *uint16
	nShow        int32
	hInstApp     windows.Handle
	lpIDList     uintptr
	lpClass      *uint16
	hkeyClass    windows.Handle
	dwHotKey     uint32
	hIcon        windows.Handle
	hProcess     windows.Handle
}

const seeMaskNoCloseProcess = 0x00000040

var procShellExecuteEx = windows.NewLazySystemDLL("shell32.dll").NewProc("ShellExecuteExW")

// runElevatedWait starts exe with the admin prompt and waits for its exit code.
func runElevatedWait(exe, args string) (uint32, error) {
	verb, _ := windows.UTF16PtrFromString("runas")
	file, _ := windows.UTF16PtrFromString(exe)
	params, _ := windows.UTF16PtrFromString(args)
	info := shellExecuteInfo{fMask: seeMaskNoCloseProcess, lpVerb: verb, lpFile: file, lpParameters: params, nShow: windows.SW_SHOWNORMAL}
	info.cbSize = uint32(unsafe.Sizeof(info))
	if r, _, err := procShellExecuteEx.Call(uintptr(unsafe.Pointer(&info))); r == 0 {
		if errors.Is(err, windows.ERROR_CANCELLED) {
			return 0, ErrDeclined
		}
		return 0, fmt.Errorf("start the installer: %w", err)
	}
	defer func() { _ = windows.CloseHandle(info.hProcess) }()
	if _, err := windows.WaitForSingleObject(info.hProcess, windows.INFINITE); err != nil {
		return 0, err
	}
	var code uint32
	if err := windows.GetExitCodeProcess(info.hProcess, &code); err != nil {
		return 0, err
	}
	return code, nil
}
