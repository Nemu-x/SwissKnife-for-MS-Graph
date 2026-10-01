# Next release — known issues and planned work

Working notes for the release after 1.1.0. Not published anywhere; the release
body is `docs/releases/v<X.Y.Z>.md`.

## Carried over from the 1.1.0 checklist

The Windows updater elevation bug is fixed in 1.1.0 (`docs/releases/v1.1.0.md`);
the real UAC round-trip still has to be exercised on a per-machine install.

## Open items

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
