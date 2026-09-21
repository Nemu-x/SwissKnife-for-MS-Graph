# Code signing

Status: **prepared, not active.** `.github/workflows/release.yml` contains gated steps for Windows
Authenticode signing (Azure Trusted Signing) and macOS Developer ID signing + notarization. Each
step checks for its secrets and is skipped when they are missing, so releases keep shipping unsigned
exactly as before until the accounts below exist. Nothing else has to change in the workflow.

Independently of this, every release already ships `SHA256SUMS.txt` signed with minisign
(`MINISIGN_SECRET_KEY` / `MINISIGN_PASSWORD`, public key in `minisign.pub`).

## Secrets the workflow looks for

| Secret | Used by | Value |
| --- | --- | --- |
| `AZURE_TRUSTED_SIGNING_TENANT_ID` | windows | Entra tenant (directory) ID of the Azure subscription |
| `AZURE_TRUSTED_SIGNING_CLIENT_ID` | windows | Application (client) ID of the app registration used for signing |
| `AZURE_TRUSTED_SIGNING_CLIENT_SECRET` | windows | Client secret of that app registration — **its presence switches Windows signing on** |
| `AZURE_TRUSTED_SIGNING_ENDPOINT` | windows | Regional endpoint of the signing account, e.g. `https://weu.codesigning.azure.net/` |
| `AZURE_TRUSTED_SIGNING_ACCOUNT` | windows | Name of the Trusted Signing account |
| `AZURE_TRUSTED_SIGNING_PROFILE` | windows | Name of the certificate profile (Public Trust) |
| `MACOS_CERTIFICATE_P12` | macos | Base64 of the exported `Developer ID Application` certificate + private key (.p12) — **its presence switches macOS signing on** |
| `MACOS_CERTIFICATE_PASSWORD` | macos | Password chosen when exporting the .p12 |
| `MACOS_SIGN_IDENTITY` | macos | Optional. `Developer ID Application: <Name> (<TEAMID>)`; auto-detected from the .p12 when empty |
| `APPLE_ID` | macos | Apple ID e-mail of the developer account — **with the two below switches notarization on** |
| `APPLE_TEAM_ID` | macos | 10-character Team ID (developer.apple.com → Membership) |
| `APPLE_APP_PASSWORD` | macos | App-specific password for that Apple ID (appleid.apple.com → Sign-In and Security) |
| `WINGET_TOKEN` | winget | Not signing, listed for completeness: classic PAT for the winget-pkgs fork (`packaging/winget/README.md`) |

Signing on macOS without the three `APPLE_*` secrets is allowed (signed but not notarized: Gatekeeper
still complains, so it is only useful for testing). Notarization without the certificate is not
possible; the notarize steps require `MACOS_CERTIFICATE_P12` too.

---

## Windows — Azure Trusted Signing

Microsoft's hosted signing service (being renamed **Artifact Signing** in the Azure portal and in the
GitHub action, `azure/artifact-signing-action`; the old `azure/trusted-signing-action` name redirects
to the same code). It issues short-lived Authenticode certificates from a Microsoft CA and signs in
the cloud, so no certificate file or HSM is ever stored in GitHub. Signed binaries are recognised by
SmartScreen as coming from a validated publisher, which removes the "Unknown publisher" warning; a
brand-new publisher identity can still see a short reputation ramp-up.

**Cost:** Basic tier ≈ **US$9.99 per month** (5,000 signatures/month, 1 certificate profile),
Premium ≈ US$99.99. Requires an Azure subscription (pay-as-you-go is fine).

**Lead time:** identity validation takes **1–7 business days** as a rule. Individuals are verified
through Microsoft Entra Verified ID (government ID + a live check); organisations need a legal entity
that is verifiable in business registries and, for Public Trust, **at least three years** of
verifiable history — younger companies or a personal project should validate as an *Individual*.

### Setup

1. **Azure subscription** — create or pick one, note its Entra tenant ID.
2. **Register the resource provider** once: `az provider register --namespace Microsoft.CodeSigning`.
3. **Create a Trusted Signing account** (portal: *Trusted Signing Accounts → Create*): pick a region
   that has the service (East US, West US, West US 2, West US 3, West Central US, North Europe,
   West Europe), SKU *Basic*. The account name and region give the endpoint,
   e.g. account in West Europe → `https://weu.codesigning.azure.net/`, East US →
   `https://eus.codesigning.azure.net/` (the portal shows the exact URI on the account overview).
4. **Identity validation** (account → *Identity validations → New*): choose Individual or
   Organization, fill in the legal details, complete the verification, wait for *Completed*.
5. **Certificate profile** (account → *Certificate profiles → Create*): type **Public Trust**, select
   the completed identity validation, give it a name (e.g. `swissknife-public`). The Subject CN of
   the certificate becomes the name shown by Windows.
6. **App registration for CI**: Entra ID → *App registrations → New*, no redirect URI; create a
   **client secret** (note the expiry — 24 months max; put a reminder in the calendar). On the
   Trusted Signing account, *Access control (IAM) → Add role assignment*: role
   **Trusted Signing Certificate Profile Signer**, assign to the app registration.
7. **GitHub secrets** (Settings → Secrets and variables → Actions):
   `AZURE_TRUSTED_SIGNING_TENANT_ID`, `AZURE_TRUSTED_SIGNING_CLIENT_ID`,
   `AZURE_TRUSTED_SIGNING_CLIENT_SECRET`, `AZURE_TRUSTED_SIGNING_ENDPOINT`,
   `AZURE_TRUSTED_SIGNING_ACCOUNT`, `AZURE_TRUSTED_SIGNING_PROFILE`.
8. **Test** with a manual `workflow_dispatch` run of *Release* (builds artifacts without publishing)
   and download the `windows-amd64` artifact:
   `Get-AuthenticodeSignature .\SwissKnifeGraph-windows-amd64-installer.exe` must report `Valid`
   and the CN of the certificate profile.

### What the workflow does

In the `windows` job, after `wails build -nsis`:

1. sign `app/build/bin/SwissKnifeGraph.exe` (SHA256, RFC3161 timestamp from
   `http://timestamp.acs.microsoft.com`);
2. re-run `makensis` exactly as Wails does (`-DARG_WAILS_<ARCH>_BINARY=<signed exe> project.nsi`) so
   the installer contains the signed executable — an installer built before signing would carry the
   unsigned exe;
3. sign the resulting `SwissKnifeGraph-<arch>-installer.exe`;
4. verify both signatures with `Get-AuthenticodeSignature` before staging.

The in-app updater downloads the installer and runs it with `/S`; a signed installer also makes the
UAC prompt show the publisher name instead of "Unknown".

---

## macOS — Developer ID + notarization

**Cost:** Apple Developer Program **US$99 per year** (individual or organisation).

**Lead time:** enrolment as an individual is usually approved within **24–48 hours**; organisations
need a D-U-N-S number and legal-entity checks (1–2 weeks). Each notarization submission takes a few
minutes; the workflow waits for it.

### Setup

1. **Enrol** at <https://developer.apple.com/programs/enroll/> with the Apple ID that will own the
   certificate. Note the **Team ID** (Account → Membership details).
2. **Create the certificate** on a Mac: Xcode → Settings → Accounts → Manage Certificates → `+` →
   *Developer ID Application*. (Or: Keychain Access → Certificate Assistant → Request a Certificate
   From a Certificate Authority, upload the CSR at developer.apple.com → Certificates → `+` →
   Developer ID Application, download and double-click the `.cer`.) Only the Account Holder can
   create Developer ID certificates.
3. **Export** it from Keychain Access (select the certificate *with* its private key → Export →
   `.p12`, set a password) and base64-encode it:
   `base64 -i DeveloperID.p12 | pbcopy`.
4. **App-specific password**: <https://appleid.apple.com> → Sign-In and Security → App-Specific
   Passwords → generate one named e.g. `github-notarytool`.
5. **GitHub secrets**: `MACOS_CERTIFICATE_P12` (the base64 text), `MACOS_CERTIFICATE_PASSWORD`,
   `APPLE_ID`, `APPLE_TEAM_ID`, `APPLE_APP_PASSWORD`, optionally `MACOS_SIGN_IDENTITY`
   (`security find-identity -v -p codesigning` prints it; the workflow detects it when omitted).
6. **Test** with a manual *Release* run, download the `macos-arm64` artifact and check:
   `spctl --assess --type open --context context:primary-signature -v SwissKnifeGraph-macos-arm64.dmg`
   should print `accepted`, and after mounting `codesign -dv --verbose=2 SwissKnifeGraph.app`
   shows `Authority=Developer ID Application: …` plus `Runtime Version`.
7. **Afterwards**: remove the `caveats` block from `packaging/homebrew/swissknife-graph.rb`, and drop
   the quarantine wording from `README.md` / `wiki/Installation.md` (the DMG stops shipping
   `Fix Quarantine.command` automatically — `create-dmg.sh` gets `DMG_SIGNED=1`).

### What the workflow does

In the `macos` job, after `wails build`:

1. import the `.p12` into a temporary keychain (`$RUNNER_TEMP/signing.keychain-db`, random
   password, deleted at the end of the job);
2. `codesign --deep --force --options runtime --timestamp --sign "$MACOS_SIGN_IDENTITY"` on the
   `.app` (hardened runtime is mandatory for notarization), then `codesign --verify --deep --strict`;
3. zip the app (`ditto -c -k --keepParent`), `xcrun notarytool submit --wait`, `xcrun stapler
   staple` the app (`packaging/macos/notarize.sh`, which fails the job with the notary log on
   anything but `Accepted`);
4. build the DMG with `DMG_SIGNED=1` (no quarantine helper / notes);
5. sign the DMG, notarize and staple it too, so Gatekeeper accepts the disk image offline as well.

Certificate renewals: Developer ID Application certificates are valid for 5 years; the app-specific
password and the membership renew yearly (the membership lapsing does not invalidate already
notarized builds).

---

## Verifying a signed release

```powershell
# Windows
Get-AuthenticodeSignature .\SwissKnifeGraph-windows-amd64-installer.exe | Format-List Status, SignerCertificate
```

```bash
# macOS
spctl --assess --type open --context context:primary-signature -v SwissKnifeGraph-macos-arm64.dmg
codesign -dv --verbose=2 /Applications/SwissKnifeGraph.app
xcrun stapler validate /Applications/SwissKnifeGraph.app
```

Linux packages are not signed beyond the minisign-signed checksum file; deb/rpm repository signing
is out of scope for a GitHub-Releases-only distribution.
