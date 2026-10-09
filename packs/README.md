# Action packs

An action pack adds to SwissKnife's catalog without rebuilding the app: a
folder with a `manifest.yaml` that declares **workflows** (chains of the app's
own actions, no code) and/or **script actions** (PowerShell scripts that run in
the app's signed-in Exchange Online or Teams PowerShell session).

See [`litigation-hold/`](litigation-hold) for a complete example.

## Install

Copy the pack folder into the `actions` folder of the app's configuration
(**Settings → Action packs → Open folder**), then trust it there.

## Trust

A pack runs only when one of these holds:

- it is **signed** with minisign by a key you trust (the project's key is built
  in; add an author's public key in Settings), or
- you **reviewed and trusted** its exact contents — any later change to any
  file makes its actions unavailable until you trust it again.

Pack scripts run in an Exchange Online or Teams session signed in with your
admin connection, so read them before trusting. The app keeps them on a short
leash:

- a script may call **only the commands its pack declares** (`cmdlets`), plus a
  few that reach neither the network nor the disk (`Where-Object`,
  `ForEach-Object`, `Select-Object`, `Sort-Object`, `Group-Object`,
  `Measure-Object`, `Compare-Object`, `Write-Output`, `Write-Verbose`,
  `Write-Warning`, `Write-Error`, `Out-Null`, `Out-String`, `ConvertTo-Json`,
  `ConvertFrom-Json`, `Get-Date`, `New-TimeSpan`, `Join-String`), each named
  literally — no `& $name`, no dot-sourcing, no `-Parallel` or `-AsJob`;
- it runs in **ConstrainedLanguage**: no .NET method calls, no `Add-Type`, no
  `[pscustomobject]` (return hashtables, `@{ ... }`, instead);
- commands that run code, start processes or jobs, or load modules
  (`Invoke-Expression`, `Invoke-Command`, `Start-Process`, `Import-Module`, …)
  can never be declared; commands that reach the network or files
  (`Invoke-RestMethod`, `Export-Csv`, `Remove-Item`, …) can, and Settings points
  them out before you trust the pack;
- they run in their own PowerShell process, apart from the built-in actions;
- every run (a preview and a "read" included) is blocked in read-only mode and
  written to the audit log;
- profile limits treat every pack action as at least a change, and a profile
  limited to groups runs no packs at all;
- a signed pack can still be turned off in Settings.

Changes to a pack folder take effect when the packs are reloaded (opening
Settings → Action packs, *Reload*, or restarting the app).

## Two kinds of packs

- **Workflow packs** chain actions the app already has — no code at all. Each
  step is planned by its action, all of it shown as one preview, confirmed
  once and written to one journal entry; profile limits apply to every step
  (also at apply). Steps are planned against the state before the run, so do
  not change the same thing twice in one workflow. On-prem AD actions and
  actions that take a password cannot be steps.
  This is the kind to write first. See [`security-basics/`](security-basics).
- **Script packs** add new actions written in PowerShell (below). Use them for
  what no built-in action does.

## Workflows

```yaml
workflows:
  - id: compromisedUser
    page: security
    label: { en: Compromised account response }
    confirmField: user            # needed when any step is destructive
    fields:
      - { name: user, kind: user, required: true }
      - { name: resetMfa, kind: choice, options: ["yes", "no"], default: "yes" }
    steps:
      - capability: entra.user.signIn
        with: { user: "{{user}}", state: blocked }
      - capability: entra.user.resetMfa
        with: { user: "{{user}}" }
        when: { input: resetMfa, equals: "yes" }   # optional step
      - capability: exchange.inboxRule.disable
        with: { user: "{{user}}" }
        onError: continue                          # default: stop
```

- A step names a built-in action by its **capability** — the stable public name
  (`exchange.mailbox.fullAccess`), listed under Settings → Capabilities — or by
  its catalog id with `action:` (`SwissKnifeGraph action list` lists every
  action and its inputs). Prefer capabilities in packs you publish: they are
  never renamed, and the app picks the implementation (Exchange REST,
  PowerShell here, or the paired worker) at run time.
- `with` gives the action's inputs; `{{name}}` takes a field of the workflow.
- A step that fails stops the run unless it says `onError: continue`.
- The workflow is as dangerous as its most dangerous step; a destructive one
  names `confirmField` (a user, group or text field), and every destructive
  step must act on exactly that field — the operator never confirms one name
  while another is changed. A step's own stronger confirmation (a purge's
  "sender N") is asked for too.

## Manifest

| Key | Meaning |
| --- | --- |
| `name` | lowercase letters, digits, dashes; unique |
| `actions[].id` | unique within the pack; the catalog id is `pack.<name>.<id>` |
| `page` | `users`, `groups`, `teams`, `mail`, `security`, `reports` or `intune` |
| `danger` | `read`, `write` or `destructive` (destructive needs `confirmField`) |
| `module` | `exo` or `teams` — which PowerShell session the script runs in |
| `script` | a `.ps1` file inside the pack |
| `cmdlets` | the commands the script calls (1–40, `Verb-Noun`); the host refuses anything else |
| `permissions` | pack level: the roles or scopes the pack needs, in words — shown before it is trusted |
| `label`, `hint` | by language (`en` required, `ru` optional) |
| `fields` | `name`, `kind` (`text`, `choice`, `user`, `group`), `required`, `options`, `default`, `label` |
| `columns` | read actions: column order |

## Script contract

`param($Mode, $Inputs, $Change)`

- `read` — return hashtables (`@{ name = ...; address = ... }`); their keys are
  the table's columns.
- `plan` — must change nothing (the app cannot enforce this — it is why a
  pack needs your trust); return `{target, field, op, before, after, ref}`
  objects (`op`: `set`, `add`, `remove` or `none`). The app previews them.
- `apply` — called once per planned change with it as `$Change`.

## Signing (authors)

```
SwissKnifeGraph pack digest ./my-pack     # writes my-pack/pack.digest
minisign -Sm ./my-pack/pack.digest         # writes pack.digest.minisig
```

Publish your minisign public key so users can add it: they give it a
publisher name, and packs it signs show "Signed by <that name>" (no key can be
named after the project's own).
