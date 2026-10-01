# winget (Windows Package Manager)

Package identifier: **`Nemu-x.SwissKnifeGraph`** → `winget install Nemu-x.SwissKnifeGraph`

Once the package exists in [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs), the
`winget` job in `.github/workflows/release.yml` opens the update PR for every new tag automatically
(via [vedantmgoyal9/winget-releaser](https://github.com/vedantmgoyal9/winget-releaser)). The job is
skipped while the `WINGET_TOKEN` secret is missing, so releases keep working without it.

The releaser **only updates packages that already exist**. The first version has to be submitted by
hand — steps 1 and 2 below are one-time work.

## 1. One-time: fork + token

1. Fork <https://github.com/microsoft/winget-pkgs> into the **Nemu-x** account. The fork must be
   named `winget-pkgs` (the releaser looks for `<repository owner>/winget-pkgs`; a fork under another
   account would need the `fork-user` input).
2. Create a **classic** personal access token (fine-grained tokens are not supported) at
   <https://github.com/settings/tokens/new> with the scopes `public_repo` and `workflow` (the second
   one lets the releaser fast-forward the fork from upstream; without it the sync fails
   intermittently with `does not have the correct permissions to execute UpdateRef`).
3. Add it to the repository as the secret **`WINGET_TOKEN`**
   (Settings → Secrets and variables → Actions → New repository secret).

## 2. One-time: first submission (1.0.0)

Ready manifests live in `packaging/winget/manifests/n/Nemu-x/SwissKnifeGraph/1.0.0/`. Only the two
`InstallerSha256` values are placeholders (`000…`, marked `# TODO`).

```powershell
# a) Install the tooling (Windows 10/11)
winget install Microsoft.WingetCreate

# b) Fill in the hashes from the release checksums
Invoke-RestMethod https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/download/v1.0.0/SHA256SUMS.txt
# copy the hashes for SwissKnifeGraph-windows-amd64-installer.exe and -arm64-installer.exe
# into Nemu-x.SwissKnifeGraph.installer.yaml (upper- or lowercase hex both validate)

# c) Validate + test-install locally (enabling LocalManifestFiles needs an elevated shell, once)
$dir = "packaging\winget\manifests\n\Nemu-x\SwissKnifeGraph\1.0.0"
winget validate --manifest $dir
winget settings --enable LocalManifestFiles
winget install --manifest $dir
winget uninstall Nemu-x.SwissKnifeGraph

# d) Submit: pushes a branch to the Nemu-x/winget-pkgs fork and opens the PR against microsoft/winget-pkgs
wingetcreate submit --token $env:WINGET_TOKEN $dir
```

Alternative if you would rather let the tool draft everything from the installers (it asks for
identifier, publisher, license, … interactively; answer with the values from the prepared files):

```powershell
wingetcreate new `
  https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/download/v1.0.0/SwissKnifeGraph-windows-amd64-installer.exe `
  https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/download/v1.0.0/SwissKnifeGraph-windows-arm64-installer.exe
# then: wingetcreate submit --token $env:WINGET_TOKEN <generated folder>
```

The PR goes through Microsoft's automated validation (installs the package in a VM) plus a human
moderator look; expect one to a few days for a first submission. Later versions from the workflow
are usually merged within hours.

## 3. After the first merge

Nothing. On the next `vX.Y.Z` tag the `winget` job runs after `publish`, downloads the release
assets matching `SwissKnifeGraph-windows-(amd64|arm64)-installer\.exe$`, bumps the manifests and
opens the PR from the fork. `max-versions-to-keep: 5` prunes old versions from winget-pkgs.

If the ARM64 leg of a release fails (it is best-effort), the PR will simply carry the x64 installer
only for that version.

## Manifest notes

- Schema `1.10.0`; `InstallerType: nullsoft`, `Scope: machine`, switches `/S` for both silent modes
  (the in-app updater uses the same switch).
- `ProductCode` is the NSIS uninstall registry key name: `Nemu-xSwissKnife for MS Graph`
  (Wails: `INFO_COMPANYNAME + INFO_PRODUCTNAME`). Keep `companyName` / `productName` in
  `app/wails.json` stable — changing them changes the ProductCode and breaks `winget upgrade`.
- A version bump needs the same three files under a new version folder; the releaser does that for
  you. To bump by hand:
  `wingetcreate update Nemu-x.SwissKnifeGraph --version X.Y.Z --urls <x64 url> <arm64 url> --submit --token …`.
