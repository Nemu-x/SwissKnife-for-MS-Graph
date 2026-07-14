# Releasing SwissKnife for MS Graph

One-time setup, then every release is a single git tag.

## 1. One-time GitHub setup

### Pages (download site)
Repo **Settings → Pages → Source = "GitHub Actions"**. The `pages.yml` workflow
regenerates the landing page from releases and deploys it to
`https://nemu-x.github.io/SwissKnife-for-MS-Graph/`.

### Secrets (Settings → Secrets and variables → Actions → New repository secret)

| Secret | Purpose | Required? |
| --- | --- | --- |
| `MINISIGN_SECRET_KEY` | Signs `SHA256SUMS.txt` so downloads can be verified | Optional (unsigned if absent) |
| `MINISIGN_PASSWORD` | Password for the minisign key (empty string if the key has none) | With the above |
| `AUR_SSH_PRIVATE_KEY` | SSH key registered on your AUR account, publishes `swissknife-graph-bin` | Optional (AUR skipped if absent) |

### Generate the minisign key (local, once)

```bash
minisign -G -p minisign.pub -s minisign.key      # prompts for a password
```

- Commit **`minisign.pub`** to the repo root (public — users verify with it).
- Put the **contents of `minisign.key`** into the `MINISIGN_SECRET_KEY` secret, and
  its password into `MINISIGN_PASSWORD`. Never commit `minisign.key`.

### AUR key (optional)

```bash
ssh-keygen -t ed25519 -f aur -C "aur@swissknife"   # add aur.pub to your AUR account
```

Put the contents of the private `aur` file into `AUR_SSH_PRIVATE_KEY`.

## 2. Cut a release

```bash
git tag v0.2.0
git push origin v0.2.0
```

That triggers `release.yml`:
1. builds Windows / macOS / Linux artifacts,
2. writes `SHA256SUMS.txt` and signs it (`.minisig`) if the minisign secret is set,
3. publishes a GitHub Release with generated notes,
4. then `pages.yml` refreshes the download site and publishes the AUR package.

The in-app updater compares the running version against the latest release tag.

## 3. Verify a download (users)

```bash
minisign -Vm SHA256SUMS.txt -P "$(cat minisign.pub)"
sha256sum -c SHA256SUMS.txt
```

## Screenshots for the README

Capture the running app (dashboard, users, offboarding, raw) and save them as
`docs/screenshots/dashboard.png`, `users.png`, `offboarding.png`, `raw.png`
(≈1600px wide). The README already references these paths.
