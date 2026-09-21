#!/usr/bin/env bash
# Submits a file (.zip of an .app, or a .dmg) to Apple's notary service, waits
# for the verdict and staples the ticket. Used by the macOS release job; only
# runs when the APPLE_* secrets are present (see packaging/SIGNING.md).
#
# Usage: notarize.sh <path-to-.zip-or-.dmg> [path-to-staple]
#   The second argument defaults to the first. For an .app pass the zip as the
#   first argument and the .app bundle as the second (a zip cannot be stapled).
#
# Env: APPLE_ID, APPLE_TEAM_ID, APPLE_APP_PASSWORD (app-specific password).
set -euo pipefail

FILE="${1:?usage: notarize.sh <zip-or-dmg> [staple-target]}"
STAPLE="${2:-$FILE}"
: "${APPLE_ID:?APPLE_ID not set}" "${APPLE_TEAM_ID:?APPLE_TEAM_ID not set}" "${APPLE_APP_PASSWORD:?APPLE_APP_PASSWORD not set}"

echo "Submitting $(basename "$FILE") for notarization…"
OUT="$(xcrun notarytool submit "$FILE" \
  --apple-id "$APPLE_ID" --team-id "$APPLE_TEAM_ID" --password "$APPLE_APP_PASSWORD" \
  --wait --timeout 45m --output-format json)"
echo "$OUT"
STATUS="$(echo "$OUT" | /usr/bin/python3 -c 'import json,sys; print(json.load(sys.stdin).get("status",""))')"
ID="$(echo "$OUT" | /usr/bin/python3 -c 'import json,sys; print(json.load(sys.stdin).get("id",""))')"

if [ "$STATUS" != "Accepted" ]; then
  echo "::error::Notarization of $(basename "$FILE") ended with status '$STATUS' (submission $ID). Log:"
  if [ -n "$ID" ]; then
    xcrun notarytool log "$ID" --apple-id "$APPLE_ID" --team-id "$APPLE_TEAM_ID" --password "$APPLE_APP_PASSWORD" || true
  fi
  exit 1
fi

xcrun stapler staple "$STAPLE"
xcrun stapler validate "$STAPLE"
echo "Notarized and stapled: $STAPLE"
