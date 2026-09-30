#!/usr/bin/env bash
# Builds a self-contained AppImage of SwissKnife for MS Graph.
#
# Usage: packaging/linux/build-appimage.sh <x86_64|aarch64> <version> <out-dir>
#   Expects the Wails binary at app/build/bin/SwissKnifeGraph and must run on a
#   Debian/Ubuntu host that has libgtk-3 and libwebkit2gtk-4.1 installed — that
#   is where the bundled copies come from. Run it from the repository root.
#
# Why this is not just "appimagetool AppDir": the app links WebKitGTK, and an
# AppImage that relies on the host's WebKit is not self-contained (the AppImage
# catalog test rejects it, and users without webkit2gtk-4.1 get "error while
# loading shared libraries"). Bundling WebKitGTK has two traps, both handled here
# the way Tauri does it:
#   1. WebKit spawns helper processes (WebKitNetworkProcess, WebKitWebProcess,
#      WebKitGPUProcess) and loads an injected bundle from an absolute path
#      compiled into libwebkit2gtk-4.1.so.0. The helpers are copied into the
#      AppDir keeping that path under usr/, and the string in the library is
#      patched to a relative path of the SAME length; AppRun chdir's into usr/
#      so it resolves to the bundled files on any distro.
#   2. GTK needs its schemas, pixbuf loaders, IM modules and GIO modules (TLS).
#      linuxdeploy-plugin-gtk (Tauri's fork, which also bundles GIO modules)
#      collects them and writes an apprun-hooks/ script that AppRun sources.
set -euo pipefail

ARCH="${1:?arch (x86_64|aarch64)}"
VERSION="${2:?version, e.g. 1.1.0}"
OUT_DIR="${3:?output directory}"

BIN=app/build/bin/SwissKnifeGraph
[ -x "$BIN" ] || { echo "::error::$BIN not found — run wails build first"; exit 1; }
LIBDIR="$(dpkg-architecture -qDEB_HOST_MULTIARCH)" # x86_64-linux-gnu / aarch64-linux-gnu
WEBKIT_LIBEXEC="/usr/lib/$LIBDIR/webkit2gtk-4.1"
[ -d "$WEBKIT_LIBEXEC" ] || { echo "::error::$WEBKIT_LIBEXEC missing — install libwebkit2gtk-4.1-0"; exit 1; }

WORK="$(mktemp -d)"
APPDIR="$WORK/AppDir"
TOOLS="$WORK/tools"
mkdir -p "$APPDIR/usr/share/applications" "$APPDIR/usr/share/icons/hicolor/512x512/apps" \
  "$APPDIR/usr/share/metainfo" "$TOOLS" "$OUT_DIR"
export APPIMAGE_EXTRACT_AND_RUN=1 # the tools are AppImages themselves; no FUSE in containers

# --- 1. Application files -------------------------------------------------
# The binary is handed to linuxdeploy under its final name (Exec= in the
# .desktop file); linuxdeploy copies it into usr/bin and sets its rpath.
cp "$BIN" "$WORK/swissknife-graph"
cp packaging/linux/swissknife-graph.desktop "$APPDIR/usr/share/applications/"
cp app/build/appicon.png "$APPDIR/usr/share/icons/hicolor/512x512/apps/swissknife-graph.png"
sed -e "s/@VERSION@/$VERSION/" -e "s/@DATE@/$(date -u +%Y-%m-%d)/" \
  packaging/linux/swissknife-graph.metainfo.xml > "$APPDIR/usr/share/metainfo/swissknife-graph.appdata.xml"
# Our own AppRun: linuxdeploy leaves an existing AppRun alone (and would
# otherwise symlink the binary, skipping the GTK hooks and the chdir).
install -m 0755 packaging/linux/AppRun "$APPDIR/AppRun"

# --- 2. WebKit helper processes + injected bundle, path preserved under usr/ -
( cd "$APPDIR" && cp -a --parents "$WEBKIT_LIBEXEC" . )
HELPERS=()
while IFS= read -r -d '' f; do HELPERS+=("$f"); done \
  < <(find "$APPDIR$WEBKIT_LIBEXEC" -type f \( -perm -u+x -o -name '*.so' \) -print0)
[ "${#HELPERS[@]}" -gt 0 ] || { echo "::error::no WebKit helpers found in $WEBKIT_LIBEXEC"; exit 1; }
echo "WebKit helpers: ${HELPERS[*]#"$APPDIR"}"

# --- 3. linuxdeploy + Tauri's GTK plugin: libraries, schemas, loaders, GIO ---
curl -fsSL -o "$TOOLS/linuxdeploy" \
  "https://github.com/linuxdeploy/linuxdeploy/releases/download/continuous/linuxdeploy-$ARCH.AppImage"
curl -fsSL -o "$TOOLS/linuxdeploy-plugin-gtk.sh" \
  "https://raw.githubusercontent.com/tauri-apps/linuxdeploy-plugin-gtk/master/linuxdeploy-plugin-gtk.sh"
chmod +x "$TOOLS/linuxdeploy" "$TOOLS/linuxdeploy-plugin-gtk.sh"
DEPS_ARGS=()
for h in "${HELPERS[@]}"; do DEPS_ARGS+=(--deploy-deps-only "$h"); done
# Helpers: dependencies only — they must stay where WebKit expects them.
PATH="$TOOLS:$PATH" DEPLOY_GTK_VERSION=3 DISABLE_COPYRIGHT_FILES_DEPLOYMENT=1 \
  "$TOOLS/linuxdeploy" --appdir "$APPDIR" \
    --executable "$WORK/swissknife-graph" \
    "${DEPS_ARGS[@]}" \
    --plugin gtk
# The helpers live three levels below usr/lib (lib/<triplet>/webkit2gtk-4.1/);
# linuxdeploy only fixed the rpath of what it copied itself.
for h in "${HELPERS[@]}"; do
  rel="$(realpath --relative-to="$(dirname "$h")" "$APPDIR/usr/lib")"
  patchelf --set-rpath "\$ORIGIN/$rel" "$h"
done
[ -x "$APPDIR/usr/bin/swissknife-graph" ] || { echo "::error::linuxdeploy did not place usr/bin/swissknife-graph"; exit 1; }
[ -f "$APPDIR/AppRun" ] && [ ! -L "$APPDIR/AppRun" ] || { echo "::error::linuxdeploy replaced AppRun"; exit 1; }
[ -d "$APPDIR/apprun-hooks" ] || { echo "::error::linuxdeploy-plugin-gtk installed no apprun hook"; exit 1; }
# Top-level desktop file, icon and .DirIcon are what appimagetool reads.
ln -sf usr/share/applications/swissknife-graph.desktop "$APPDIR/swissknife-graph.desktop"
ln -sf usr/share/icons/hicolor/512x512/apps/swissknife-graph.png "$APPDIR/swissknife-graph.png"
ln -sf swissknife-graph.png "$APPDIR/.DirIcon"

# --- 4. Patch the compiled-in libexec path (same length, fail-closed) --------
OLD="$WEBKIT_LIBEXEC"                         # /usr/lib/<triplet>/webkit2gtk-4.1
NEW="././/lib/$LIBDIR/webkit2gtk-4.1"          # resolves from cwd=usr/ (AppRun)
[ "${#OLD}" -eq "${#NEW}" ] || { echo "::error::patch strings differ in length"; exit 1; }
export OLD NEW
patched=0
for lib in "$APPDIR"/usr/lib/libwebkit2gtk-4.1.so*; do
  [ -f "$lib" ] && [ ! -L "$lib" ] || continue
  before="$(perl -0777 -ne 'my $c = () = /\Q$ENV{OLD}\E/g; print $c' "$lib")"
  [ "$before" -gt 0 ] || { echo "::error::$lib does not contain $OLD — WebKit layout changed, refusing to ship"; exit 1; }
  size_before="$(stat -c %s "$lib")"
  perl -0777 -pi -e 's/\Q$ENV{OLD}\E/$ENV{NEW}/g' "$lib"
  after="$(perl -0777 -ne 'my $c = () = /\Q$ENV{OLD}\E/g; print $c' "$lib")"
  [ "$after" -eq 0 ] || { echo "::error::$lib still contains $OLD after patching"; exit 1; }
  [ "$(stat -c %s "$lib")" -eq "$size_before" ] || { echo "::error::$lib changed size while patching"; exit 1; }
  echo "Patched $before occurrence(s) of $OLD in ${lib#"$APPDIR"/}"
  patched=$((patched + 1))
done
[ "$patched" -gt 0 ] || { echo "::error::libwebkit2gtk-4.1.so was not bundled"; exit 1; }

# --- 5. Pack --------------------------------------------------------------
curl -fsSL -o "$TOOLS/appimagetool" \
  "https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-$ARCH.AppImage"
chmod +x "$TOOLS/appimagetool"
OUT="$OUT_DIR/SwissKnifeGraph-$ARCH.AppImage" # no "linux" in the name: every AppImage is for Linux
ARCH="$ARCH" "$TOOLS/appimagetool" --no-appstream "$APPDIR" "$OUT"
chmod +x "$OUT"
echo "Built $OUT ($(du -h "$OUT" | cut -f1))"
echo "appdir=$APPDIR" >> "${GITHUB_OUTPUT:-/dev/null}"
