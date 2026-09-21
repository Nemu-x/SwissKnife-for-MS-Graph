# EWS is being switched off. If you are re-tooling anyway, here is a Graph-native desktop client for the ticket queue

*Draft for r/sysadmin / r/Office365 / Microsoft Tech Community. Maintainer voice, first person.*

On 1 October 2026 Exchange Web Services stops being supported in Exchange
Online: tenants that have not put their EWS apps on an allow list get
`EWSEnabled` flipped to false, and in April 2027 the protocol is removed for
good. Earlier this year, on 8 April, the Reporting Web Service behind
`Get-MessageTrace` / `Get-MessageTraceDetail` started its own deprecation, with
message trace moving to Graph. Neither change is news, but together they mean
that a lot of admins are re-examining what talks to their tenant right now, and
a fair number of "I'll just run this old script" habits are dying.

I maintain a small open-source desktop tool that lives entirely on Microsoft
Graph. Here is what it does for the admin who does not want to script every
ticket, and what it cannot do — the gaps are Graph's gaps and you will hit
them whatever tool you pick.

## What it is

SwissKnife for MS Graph is a Windows / macOS / Linux desktop client (Go
backend, React front end, Wails) that authenticates against your own app
registration, either app-only with a client secret kept in the OS keychain, or
delegated through device code. The token never reaches the web layer. Every
write goes through a typed confirmation and a local audit log; there is a
global read-only switch for when you just want to look. MIT licensed.

The design decision that matters is that the UI is organised around **tasks,
not endpoints**. Each page is a grid of the jobs it can do, named the way a
ticket names them: "add someone to a private channel", "why can this person
not sign in?", "give Bell the same access Whitfield has". Press Ctrl+K and type
a few words (English or Russian, both work at once) and you land on the form
that does it. Next to every form is a short note with the Graph rules that
will bite: a private channel only takes existing team members, dynamic and
Exchange-managed groups cannot be changed through Graph, removing the last
license deletes the mailbox after roughly 30 days. The raw table / JSON / tree
views are still there, one click away; they are just not the first thing you
see.

A few of the tasks:

- **Access mirror.** Reads two people's groups, admin roles, teams, private
  channels and licenses, shows the diff (including what only the *target*
  has), and copies what is missing. Additive only, typed confirmation,
  journalled.
- **Offboarding playbook.** Block, revoke sessions, strip MFA methods,
  auto-reply, forwarding via an inbox rule, hide from GAL, hand owned groups
  and teams to someone else, cancel the meetings the leaver organises, remove
  from groups, retire or wipe Intune devices, delete registered devices,
  server-side copy of the OneDrive into someone else's OneDrive, remove
  licenses, delete. Each step reports separately; a 403 tells you which
  permission is missing instead of just "Forbidden".
- **Audit that answers a question.** One user's sign-ins, failures first, with
  the Entra error behind each row, over a window you choose.
- **Conditional Access as facts.** Who a policy applies to, who is excluded,
  which apps, what it demands. JSON folded away.
- **Message trace** (next release). "Where did the email go?" against the new
  `/admin/exchange/tracing/messageTraces` endpoint, with the one-time
  Transport Data Platform service principal set up for you, because that
  part is not obvious from the docs.
- Plus the ordinary things: users, licenses, groups, Teams, chats, mail,
  files, Intune, BitLocker keys, secret expiry, usage reports, storage
  cleanup, and a raw Graph playground with history.

## What it cannot do, and why

These are limits of Graph, not of the app, and I would rather you knew them
before you install anything.

- **Convert a mailbox to shared.** There is no Graph API for it. The
  offboarding playbook is therefore designed as two runs: everything except
  licenses, then you convert the mailbox by hand in the Exchange admin center,
  then licenses only, with a pre-flight step that checks the mailbox type
  before removing anything.
- **Teams chat backup needs protected-API approval.** Reading chats app-only
  requires `Chat.Read.All`, which Microsoft must approve per app id beyond
  admin consent. Until that lands for your tenant, the step fails with 403,
  and the app says so.
- **Hide from GAL on AD-synced users** fails, because the attribute is mastered
  on-premises. The app reports Graph's error; fix it in AD.
- **The macOS build is unsigned** for now. The DMG ships a "Fix Quarantine"
  script and the `xattr` one-liner. Windows has minisign-signed checksums but
  no Authenticode, so SmartScreen warns on first run. Both are on the list;
  neither is free.

## Trying it

- Windows: `winget install Nemu-x.SwissKnifeGraph` (pending in the community
  repo at the time of writing; check the README), or download the installer /
  portable exe from the releases page.
- Arch: `yay -S swissknife-graph-bin`.
- macOS and Linux (deb / rpm / AppImage / tar.gz): direct download from
  https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases. Checksums are
  signed with minisign; the public key is in the repo.

You need an app registration; the wiki has the permission matrix per feature
and a minimal set to start with. Turn on read-only mode first, click around,
then grant write permissions for the tasks you actually use.

## What I would like from you

This is a one-person project and I am the only tenant I test against
regularly. If you try it, I want to hear:

- which ticket you reached for it and whether the tile said what you expected;
- every 400/403 whose message did not tell you what to fix;
- what you still have to open PowerShell for.

Issues and feature requests: https://github.com/Nemu-x/SwissKnife-for-MS-Graph/issues.
Discussions are open on the repo too. I read everything.

## Sources

- Microsoft Graph — what's new: https://learn.microsoft.com/en-us/graph/whats-new-overview
- Deprecation of Exchange Web Services in Exchange Online: https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online
- Message Trace support using Graph API is now in public preview: https://techcommunity.microsoft.com/blog/exchange/message-trace-support-using-graph-api-is-now-in-public-preview/4488587
- General availability of the mailbox import and export Microsoft Graph APIs: https://devblogs.microsoft.com/microsoft365dev/announcing-general-availability-of-the-mailbox-import-and-export-microsoft-graph-apis/
- Project: https://github.com/Nemu-x/SwissKnife-for-MS-Graph
