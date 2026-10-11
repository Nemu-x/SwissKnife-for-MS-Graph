# Next release — known issues and planned work

Working notes for the release after 2.0.0. Not published anywhere; the release
body is `docs/releases/v<X.Y.Z>.md`.

## Open items

- PowerShell throttling: back off on Exchange/Teams "server busy" / budget
  errors for reads and idempotent writes (today only auth errors are retried).
- Quarantine list reads one page of 100: page through with a cap and flag
  truncation.
- GitHub Actions: move to the Node 24 majors; pin `nfpm` and the Ubuntu runner;
  plan for the end of the Intel macOS runner.
- Code signing (Windows Authenticode, macOS Developer ID + notarization): the
  release workflow has gated steps — see `packaging/SIGNING.md`.
- Scoop bucket and Homebrew tap repositories (`packaging/README.md`); add them
  back to the README install table once they exist.
- winget: after microsoft/winget-pkgs#438961 is merged, drop
  `continue-on-error` from the release `winget` job.
- Tenant smoke (R.2): add the `SMOKE_*` secrets and extend with Exchange Admin
  API and PowerShell reads.
- `Chat.Read.All` protected-API approval (aka.ms/teamsgraph/requestaccess) —
  until it lands, the Teams chat backup step stays blocked.
- The real UAC round-trip of the Windows updater on a per-machine install.
