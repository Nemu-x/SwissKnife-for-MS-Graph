# Releasing SwissKnife for MS Graph

One-time setup, then every release is two version bumps, notes, and a tag on `main`.

## 1. One-time GitHub setup

### Pages (download site)
Repo **Settings → Pages → Source = "GitHub Actions"**. The `pages.yml` workflow
regenerates the landing page from releases, deploys it to
`https://nemu-x.github.io/SwissKnife-for-MS-Graph/`, and publishes the AUR package.
It runs when the Release workflow **succeeds**.

### Secrets (Settings → Secrets and variables → Actions → New repository secret)

| Secret | Purpose | Required? |
| --- | --- | --- |
| `MINISIGN_SECRET_KEY` | Signs `SHA256SUMS.txt` with the **release key** (`minisign.pub`); the in-app updater refuses an installer whose checksum is not signed by it | Yes |
| `MINISIGN_PASSWORD` | Password for that key — leave the secret out if the key has none | With the above |
| `AUR_SSH_PRIVATE_KEY` | SSH key registered on the AUR account, publishes `swissknife-graph-bin` | Optional (AUR skipped if absent) |
| `WINGET_TOKEN` | Opens the winget-pkgs update PR for `Nemu-x.SwissKnifeGraph` | Optional (skipped if absent) |
| `AZURE_*`, `MACOS_*` / `APPLE_*` | Windows Authenticode / macOS Developer ID + notarization | Optional — see [`packaging/SIGNING.md`](../packaging/SIGNING.md) |

Two minisign keys exist and never overlap:

- the **release key** (`minisign.pub` here, secret above) signs release checksums;
- the **pack key** (`pack.pub` in [SwissKnife-Action-Hub](https://github.com/Nemu-x/SwissKnife-Action-Hub),
  secret `PACK_SIGNING_KEY` in that repo) signs Action Hub packs and is built into the app
  (`app/internal/packs/packs.go`, `BuiltinKey`).

Keep an offline backup of both secret keys: a GitHub secret cannot be read back.

### Generate a minisign key (local, once)

```bash
minisign -G -p minisign.pub -s minisign.key      # prompts for a password
```

Commit the public key; put the whole contents of the secret key file (both lines)
into the secret. Never commit the secret key.

## 2. Cut a release

1. On a branch from `dev`, bump the version in **both** places the `version-check`
   job compares with the tag:
   - `app/app.go` — `baseVersion`
   - `app/wails.json` — `info.productVersion`
2. Write the release notes as `docs/releases/v<X.Y.Z>.md` — this file becomes the
   release body (without it the release falls back to generated notes).
3. Merge to `dev`, then `dev` → `main` (PR).
4. Before tagging a big release, dry-run the build: `gh workflow run release.yml --ref dev`
   and check every platform leg is green.
5. Tag **on `main`**:

   ```bash
   git checkout main && git pull
   git tag v2.0.0
   git push origin v2.0.0
   ```

The tag triggers `release.yml`:

1. `version-check` — the tag must equal `baseVersion` and `productVersion`;
2. builds Windows (installer + portable, amd64/arm64), macOS (arm64/intel DMG) and
   Linux (tar.gz, deb, rpm, AppImage + `.zsync`, amd64/arm64);
3. writes `SHA256SUMS.txt`, signs it (`.minisig`), publishes the GitHub Release with
   the curated notes;
4. `winget` opens the winget-pkgs update PR (it cannot update a package that is not
   in winget-pkgs yet, so it is allowed to fail without failing the release);
5. `pages.yml` then refreshes the download site and publishes AUR — check that it ran.

The in-app updater compares the running version with the latest release tag and
verifies the installer against the signed checksums before running it.

## 3. Verify a download (users)

```bash
minisign -Vm SHA256SUMS.txt -P "$(cat minisign.pub)"
sha256sum -c SHA256SUMS.txt
```

## Screenshots for the README

`docs/screenshots/*.png` are captured from the stubbed GUI:
`cd app/frontend && npx vite build && SHOTS=1 npx playwright test e2e/screenshots.spec.ts`.
