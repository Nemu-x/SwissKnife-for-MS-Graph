#!/usr/bin/env bash
# Smoke-tests a built AppImage the way the AppImage catalog does: on a machine
# WITHOUT webkit2gtk, it must load (CLI `version`), and the GUI must stay alive
# past WebKit's helper-process spawn instead of crashing at startup.
#
# Usage: packaging/linux/test-appimage.sh <path/to/App.AppImage> [appdir]
# Meant for the release container (Ubuntu 22.04, root): it REMOVES the host
# webkit2gtk packages first. Do not run it on a workstation.
set -euo pipefail

AI="${1:?AppImage path}"
APPDIR="${2:-}"
[ -x "$AI" ] || { echo "::error::$AI missing or not executable"; exit 1; }
export APPIMAGE_EXTRACT_AND_RUN=1

echo "== removing host WebKitGTK so only the bundled copy can satisfy the loader"
apt-get remove -y --purge 'libwebkit2gtk-4.1-*' 'libjavascriptcoregtk-4.1-*' >/dev/null
! ldconfig -p | grep -q libwebkit2gtk || { echo "::error::host webkit still present"; exit 1; }

echo "== headless: the binary and every bundled library must load"
out="$("$AI" version 2>&1)" || { echo "$out"; echo "::error::AppImage failed to start headless"; exit 1; }
echo "$out" | grep -q SwissKnifeGraph || { echo "$out"; echo "::error::unexpected version output"; exit 1; }
echo "$out"

echo "== GUI under Xvfb: must still be running after 20 s (timeout exit 124)"
set +e
xvfb-run -a --server-args='-screen 0 1280x800x24' timeout 20s "$AI" >gui.log 2>&1
rc=$?
set -e
cat gui.log || true
[ "$rc" -eq 124 ] || { echo "::error::GUI exited early with code $rc"; exit 1; }
if grep -Eiq 'error while loading shared|Unable to spawn|Failed to fully launch|cannot open shared object' gui.log; then
  echo "::error::startup log shows a bundling problem"; exit 1
fi

if [ -n "$APPDIR" ] && command -v desktop-file-validate >/dev/null; then
  echo "== desktop file"
  desktop-file-validate "$APPDIR/usr/share/applications/swissknife-graph.desktop"
fi
if [ -n "$APPDIR" ] && command -v appstreamcli >/dev/null; then
  echo "== AppStream metainfo (warnings allowed, errors fail)"
  appstreamcli validate --no-net "$APPDIR"/usr/share/metainfo/*.xml
fi
echo "AppImage smoke test passed."
