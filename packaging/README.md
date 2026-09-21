# Packaging & distribution

Everything the release pipeline (`.github/workflows/release.yml`, `pages.yml`) needs to turn a
`vX.Y.Z` tag into installable packages, plus the manifests for third-party package managers.

| Channel | Status | Where |
| --- | --- | --- |
| GitHub Releases (Windows installer + portable, macOS DMGs, Linux AppImage/deb/rpm/tar.gz) | live, automatic | `release.yml`, `macos/create-dmg.sh`, `nfpm.yaml`, `linux/` |
| Download site (GitHub Pages) | live, automatic | `site/generate.py`, `pages.yml` |
| AUR `swissknife-graph-bin` | live, automatic | `aur/PKGBUILD.template`, `pages.yml` (needs `AUR_SSH_PRIVATE_KEY`) |
| winget `Nemu-x.SwissKnifeGraph` | **pending first manual submission**, then automatic | [`winget/`](winget/README.md) |
| Scoop `swissknife-graph` | **pending**: needs a bucket repo | [`scoop/`](#scoop) |
| Homebrew cask `swissknife-graph` | **pending**: needs a tap repo | [`homebrew/`](#homebrew-cask) |
| Chocolatey | not planned | [below](#chocolatey-not-planned) |
| Code signing (Windows Authenticode, macOS notarization) | **prepared, off until secrets exist** | [`SIGNING.md`](SIGNING.md) |

Every optional channel is gated on a secret: when the secret is missing the corresponding job or
step is skipped and the release is exactly what it is today.

## winget

See [`winget/README.md`](winget/README.md): fork `microsoft/winget-pkgs`, create the `WINGET_TOKEN`
PAT, submit 1.0.0 once by hand with `wingetcreate`, and the `winget` job takes over from the next
tag.

## Scoop

`scoop/swissknife-graph.json` installs the **portable** exe (renamed to `SwissKnifeGraph.exe`, shim
`swissknife-graph`, Start-menu shortcut) for x64 and ARM64. `checkver: github` tracks the latest
release tag and `autoupdate` pulls both hashes straight from the release's `SHA256SUMS.txt`, so a
bucket with Scoop's standard excavator workflow updates itself.

Publishing in a personal bucket (recommended, no review queue):

```powershell
# 1. Create the repo Nemu-x/scoop-bucket from https://github.com/ScoopInstaller/BucketTemplate
#    ("Use this template"). It ships bin/checkver.ps1, bin/checkhashes.ps1 and the
#    .github/workflows/excavator.yml auto-update job.
# 2. Copy the manifest and fill the hashes:
Copy-Item packaging\scoop\swissknife-graph.json <scoop-bucket>\bucket\
cd <scoop-bucket>
.\bin\checkhashes.ps1 swissknife-graph -Update     # replaces the TODO hashes from SHA256SUMS.txt
.\bin\checkver.ps1 swissknife-graph                # should print the current version, no update
# 3. Commit + push. Users then run:
scoop bucket add nemu-x https://github.com/Nemu-x/scoop-bucket
scoop install swissknife-graph
```

Submitting to the official `extras` bucket instead: open a PR to
<https://github.com/ScoopInstaller/Extras> adding `bucket/swissknife-graph.json` (same file, hashes
filled). Their CI runs `checkver`/`checkhashes` and a maintainer reviews it; GUI apps with a stable
GitHub release cadence are routinely accepted. Do this once winget/Homebrew are in place — the
personal bucket keeps working either way.

## Homebrew cask

`homebrew/swissknife-graph.rb` is a cask for a personal tap: separate DMG URLs and `sha256` values
for Apple Silicon (`arm64`) and Intel (`intel`), `app "SwissKnifeGraph.app"`, a `livecheck` on the
GitHub releases, `zap` for the app's data directories, and a `caveats` block explaining the missing
notarization (delete it once `SIGNING.md` is done).

```bash
# 1. Create the repo Nemu-x/homebrew-tap (the "homebrew-" prefix lets users write Nemu-x/tap).
mkdir -p homebrew-tap/Casks && cp packaging/homebrew/swissknife-graph.rb homebrew-tap/Casks/
# 2. Fill the two sha256 values from the release's SHA256SUMS.txt, then check the cask:
brew tap Nemu-x/tap /path/to/homebrew-tap      # or after pushing: brew tap Nemu-x/tap
brew audit --cask --online swissknife-graph
brew style --fix homebrew-tap/Casks/swissknife-graph.rb
brew install --cask swissknife-graph           # test install / brew uninstall --zap
# 3. Commit + push. Users then run:
brew install --cask Nemu-x/tap/swissknife-graph
# Bumping later: brew bump-cask-pr --version X.Y.Z swissknife-graph  (or edit version + sha256)
```

Homebrew's main cask repository (`homebrew/cask`) requires notarized apps and a minimum popularity
(GitHub stars / downloads), so the personal tap is the realistic option until signing lands.

## Chocolatey (not planned)

Chocolatey would need its own package (a `.nuspec` plus `chocolateyinstall.ps1` calling the NSIS
installer with `/S` and the SHA256 from `SHA256SUMS.txt`) pushed to the community feed with an
API key, and every version then sits in a moderation queue where automated virus scans and a human
moderator check it, which takes days and regularly stalls on unsigned binaries and SmartScreen
reputation. It also has to be kept in sync by hand or via an extra `au`/Chocolatey-AU update job.
winget and Scoop cover the same audience on Windows with far less upkeep, so Chocolatey is skipped
for now; if demand shows up, the manifests in `winget/` contain every value the nuspec would need.

## Signing

See [`SIGNING.md`](SIGNING.md) for the Azure Trusted Signing (Windows) and Apple Developer ID +
notarization (macOS) setup, the secrets the workflow expects, costs and lead times.
