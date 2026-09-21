# Next release — known issues and planned work

Working notes for the release after 1.0.0. Not published anywhere; the release
body is `docs/releases/v<X.Y.Z>.md`.

## Fixed

### In-app update fails on Windows: "requires elevation"

Reported on the 1.0.0 build. **Update now** downloaded the installer, then died
with `fork/exec ...-installer.exe: The requested operation requires elevation`.
`UpdateService.Apply` started the per-machine NSIS installer
(`InstallDir "$PROGRAMFILES64\..."`) with `exec.Command(clean, "/S").Start()`,
which Windows refuses without elevation (ERROR_ELEVATION_REQUIRED, 740) — and
`os/exec` cannot raise a UAC prompt — so the path never worked on a normal
install.

Fix: `Apply` now launches the installer through the shell with the `runas` verb
(`windows.ShellExecute(0, "runas", exe, "/S", nil, SW_HIDE)` from
`golang.org/x/sys/windows`, behind build tags in `update_windows.go` /
`update_other.go`), keeping the temp-dir and `-installer.exe` guards; the app
quits only after a successful launch. A dismissed UAC dialog (ERROR_CANCELLED,
1223) comes back as `ErrUpdateDeclined` (OpError code `update_declined`), which
Settings shows as a friendly "cancelled" toast and keeps **Update now** usable.
Any other launch failure shows the error plus a **Show downloaded installer**
button (new `UpdateService.RevealInstaller`, `explorer.exe /select,<path>`) so
the user can run the file by hand. Settings also states up front that Windows
will ask for administrator approval once.

Rejected alternative: a per-user installer (`RequestExecutionLevel user` into
`$LOCALAPPDATA`) removes the prompt but moves the install location and breaks
upgrades from existing per-machine installs.

## Carried over from the 1.0.0 checklist

- Wiki push (grown on the post-1.0 branch): Permissions.md (rows for message trace, mailbox copy, recommendations, snapshot), new pages Task-First-UI.md, Offboarding.md, CLI.md, updated Installation.md, Home.md, _Sidebar.md — plus the older items: (new: `Group.ReadWrite.All` for ownership transfer,
  `Calendars.ReadWrite` for cancelling meetings), Teams-Notifications.md, a page
  on the task-first UI.
- `Chat.Read.All` protected-API approval (aka.ms/teamsgraph/requestaccess) —
  until it lands, the Teams chat backup step stays blocked.
- Code signing (Windows) and Developer ID signing + notarization (macOS): the
  release workflow now has gated steps for both — see `packaging/SIGNING.md`
  for the secrets to create. Same for winget (`WINGET_TOKEN` + a one-time
  manual `wingetcreate` submission, `packaging/winget/README.md`), Scoop bucket
  and Homebrew tap (`packaging/README.md`).
- Connect / Settings / Run history still carry their own headers; they stay
  form-and-list pages, but the shared header treatment is missing.
