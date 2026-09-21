# Launch posts — where and how

Drafts in this folder:

- `2026-ews-retirement-en.md` — English post (r/sysadmin, r/Office365, r/entra, Microsoft Tech Community).
- `2026-ews-retirement-ru.md` — Russian post written for Habr (not a translation).

Both are timed to the EWS retirement (1 October 2026 disablement starts, April
2027 full removal) and the Reporting Web Service message-trace deprecation
(8 April 2026). Post before 1 October while the topic is live.

## Where to post

| Venue | URL | Notes |
| --- | --- | --- |
| r/sysadmin | https://www.reddit.com/r/sysadmin/ | Self-promotion rules are strict: post as a text post, lead with the EWS angle, link the repo once at the end. No affiliate/donation links in the post body. Read the sidebar rules the day you post. |
| r/Office365 | https://www.reddit.com/r/Office365/ | Same text, smaller audience, more admins who actually run Exchange Online. |
| r/entra | https://www.reddit.com/r/entra/ | Shorter version focused on Access mirror / Conditional Access / sign-in audit. |
| Microsoft Tech Community | https://techcommunity.microsoft.com/ | Post in the Exchange or Microsoft 365 discussion space (not the official blogs). Tech Community allows longer form; keep the Sources list. |
| Habr | https://habr.com/ru/hubs/sys_admin/ | Hub «Системное администрирование»; also tag «Microsoft 365», «Microsoft Graph», «open source». Habr wants a real technical article, not a product page — the RU draft is written for that. Add 2–3 screenshots from `docs/screenshots/`. |

Optional: LinkedIn (short summary + link), Bluesky/X with a screenshot of the
palette.

## Community tool lists to submit to

Verified 2026-09 (fetched and checked):

| List | URL | How to submit |
| --- | --- | --- |
| awesome-entra (Merill Fernando) | https://github.com/merill/awesome-entra | Curated list with a **Tools** section (CLI, Web apps…). No CONTRIBUTING.md; submit a pull request adding a line under Tools. One-line description, link to the repo, no marketing. |
| merill.net | https://merill.net/ | Merill's hub page listing his own tools (Maester, cmd.ms, Graph X-Ray, Graph Permissions Explorer at https://graphpermissions.merill.net, idPowerToys…). It does **not** list third-party tools and has no submission form — do not submit here; it is useful as a reference in posts. |
| Entra.News | https://entra.news/ | Weekly newsletter (Substack, ~19k subscribers) curated by Merill Fernando. No public submission form; the practical route is to publish the post, then mention it to Merill (Bluesky/LinkedIn) or reply to the newsletter email. |
| M365 Message Center archive | https://mc.merill.net/ | Not a tool list — but the EWS message (MC1227454) and the message-trace message (MC1221939) are good citations for the posts. |

Not verified / could not find: a dedicated "community tools" directory on
merill.net other than awesome-entra. If Merill publishes one later, add it
here.

Other places worth a PR or an issue:

- https://github.com/microsoftgraph — nothing to submit, but watch the
  "Community" repos for calls for samples.
- AUR page for `swissknife-graph-bin` — keep the description current; it is
  the first thing Arch users see.

## Checklist for the maintainer

Before posting:

- [ ] 1.1.0 is tagged and the release page shows all assets (winget/scoop/brew
      lines in the posts must be true; if a manifest is still pending, keep
      the "pending" parenthesis or drop the line).
- [ ] Wiki pushed: Task-First-UI, Offboarding, Permissions (new rows).
- [ ] Screenshots in `docs/screenshots/` regenerated from the 1.1.0 bundle.
- [ ] README download table matches the release assets.
- [ ] Re-read both posts once for anything that reads as marketing; cut it.

Posting:

- [ ] Reddit: r/sysadmin first (text post), wait a day, then r/Office365 and
      r/entra with the shortened variant. Do not cross-post the same hour.
- [ ] Tech Community discussion post, with the Sources list.
- [ ] Habr, hub «Системное администрирование», with screenshots.
- [ ] PR to merill/awesome-entra (Tools section).
- [ ] Short LinkedIn/Bluesky mention linking the Habr or Reddit post, not the repo.

After posting:

- [ ] Answer every comment for the first 48 hours; turn reproducible
      complaints into GitHub issues yourself and link them back.
- [ ] Note which venue produced issues/stars in `docs/releases/NEXT.md` for the
      next launch.
