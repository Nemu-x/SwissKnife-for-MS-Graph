#!/usr/bin/env bash
# Creates a styled DMG from the .app produced by `wails build`: drag-to-/Applications
# layout, install notes, and a Gatekeeper quarantine-fix helper (the app is unsigned).
#
# Usage: create-dmg.sh <bin-dir> <out-dmg-path> [volume-name]
#   bin-dir      directory containing the built .app (e.g. app/build/bin)
#   out-dmg-path where to write the final compressed DMG
#   volume-name  mounted volume title (default: "SwissKnife for MS Graph")
#
# Optional: docs/dmg-background.png (repo root) is used as the window background
# when present; otherwise the window is plain white.
set -euo pipefail

BIN_DIR="${1:?usage: create-dmg.sh <bin-dir> <out-dmg-path> [volume-name]}"
OUT_DMG="${2:?usage: create-dmg.sh <bin-dir> <out-dmg-path> [volume-name]}"
VOLNAME="${3:-SwissKnife for MS Graph}"

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
DMG_BG="$REPO_ROOT/docs/dmg-background.png"

APP_PATH="$(find "$BIN_DIR" -maxdepth 1 -name '*.app' | head -1)"
[ -n "$APP_PATH" ] || { echo "No .app found in $BIN_DIR" >&2; exit 1; }
APP_NAME="$(basename "$APP_PATH")"

STAGE="$(mktemp -d)/dmg-stage"
TMP_DMG="${OUT_DMG%.dmg}-tmp.dmg"
mkdir -p "$STAGE"
cp -R "$APP_PATH" "$STAGE/$APP_NAME"
ln -s /Applications "$STAGE/Applications"

cat > "$STAGE/README.txt" <<EOF
SwissKnife for MS Graph — macOS install
=======================================

1) Drag "$APP_NAME" to "Applications".
2) Launch it from Applications.

If macOS blocks startup (Gatekeeper quarantine), run:
  sudo xattr -r -d com.apple.quarantine "/Applications/$APP_NAME"

You can also double-click "Fix Quarantine.command" in this DMG after drag&drop.
EOF

cat > "$STAGE/Fix Quarantine.command" <<EOF
#!/bin/bash
set -euo pipefail
APP="/Applications/$APP_NAME"
if [ ! -d "\$APP" ]; then
  osascript -e 'display dialog "Install the app to /Applications first." buttons {"OK"} default button "OK" with icon caution'
  exit 1
fi
sudo xattr -r -d com.apple.quarantine "\$APP" || true
osascript -e 'display dialog "Done. You can now open SwissKnife for MS Graph." buttons {"OK"} default button "OK"'
EOF
chmod 0755 "$STAGE/Fix Quarantine.command"

HAS_BG=false
if [ -f "$DMG_BG" ]; then
  mkdir -p "$STAGE/.background"
  cp "$DMG_BG" "$STAGE/.background/dmg-background.png"
  HAS_BG=true
fi

rm -f "$OUT_DMG" "$TMP_DMG"
hdiutil create -volname "$VOLNAME" -srcfolder "$STAGE" -ov -format UDRW "$TMP_DMG"

ATTACH_OUT="$(hdiutil attach -readwrite -noverify -noautoopen "$TMP_DMG")"
DEVICE="$(echo "$ATTACH_OUT" | awk '/^\/dev\// {print $1; exit}')"
[ -n "$DEVICE" ] || { echo "Failed to parse mounted device:"; echo "$ATTACH_OUT"; exit 1; }

if [ "$HAS_BG" = true ]; then
  BG_LINE='set background picture of opts to file ".background:dmg-background.png"'
else
  BG_LINE='set background color of opts to {65535, 65535, 65535}'
fi

osascript <<EOF
tell application "Finder"
  tell disk "$VOLNAME"
    open
    set current view of container window to icon view
    set toolbar visible of container window to false
    set statusbar visible of container window to false
    set the bounds of container window to {120, 120, 980, 650}
    set opts to the icon view options of container window
    set arrangement of opts to not arranged
    set icon size of opts to 120
    $BG_LINE
    delay 0.2
    set position of item "$APP_NAME" of container window to {230, 250}
    set position of item "Applications" of container window to {700, 250}
    try
      set position of item "README.txt" of container window to {230, 500}
    end try
    try
      set position of item "Fix Quarantine.command" of container window to {450, 500}
    end try
    close
    open
    update without registering applications
    delay 0.2
    close
  end tell
end tell
EOF

# hdiutil detach transiently fails with "Resource busy" on CI runners while
# Spotlight/Finder still hold the volume; retry, escalating to -force.
detached=false
for attempt in 1 2 3 4; do
  if [ "$attempt" -eq 1 ]; then
    hdiutil detach "$DEVICE" && detached=true && break || true
  else
    echo "detach attempt $attempt (retrying with -force)…"
    sleep 2
    hdiutil detach "$DEVICE" -force && detached=true && break || true
  fi
done
[ "$detached" = true ] || { echo "Failed to detach $DEVICE" >&2; exit 1; }

hdiutil convert "$TMP_DMG" -format UDZO -imagekey zlib-level=9 -o "$OUT_DMG"
rm -f "$TMP_DMG"
rm -rf "$STAGE"
echo "Wrote $OUT_DMG"
