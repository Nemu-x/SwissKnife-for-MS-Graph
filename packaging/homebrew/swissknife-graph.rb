# Homebrew cask for a personal tap: repo Nemu-x/homebrew-tap, file Casks/swissknife-graph.rb.
#   brew install --cask Nemu-x/tap/swissknife-graph
#
# The DMGs are unsigned/un-notarized until packaging/SIGNING.md is done; the
# caveats block tells users how to clear the Gatekeeper quarantine flag. Remove
# the block once releases are notarized.
#
# Bumping: `brew bump-cask-pr --version X.Y.Z swissknife-graph` inside the tap,
# or edit `version` and both sha256 values by hand (from the release's SHA256SUMS.txt).
cask "swissknife-graph" do
  arch arm: "arm64", intel: "intel"

  version "1.0.0"
  sha256 arm:   "TODO_SHA256_OF_SwissKnifeGraph-macos-arm64.dmg",
         intel: "TODO_SHA256_OF_SwissKnifeGraph-macos-intel.dmg"

  url "https://github.com/Nemu-x/SwissKnife-for-MS-Graph/releases/download/v#{version}/SwissKnifeGraph-macos-#{arch}.dmg",
      verified: "github.com/Nemu-x/SwissKnife-for-MS-Graph/"
  name "SwissKnife for MS Graph"
  desc "Microsoft Graph desktop client for IT admins"
  homepage "https://github.com/Nemu-x/SwissKnife-for-MS-Graph"

  livecheck do
    url :url
    strategy :github_latest
  end

  # LSMinimumSystemVersion in app/build/darwin/Info.plist is 10.13.
  depends_on macos: ">= :high_sierra"

  app "SwissKnifeGraph.app"

  # Bundle id is com.wails.SwissKnifeGraph (Wails default). Client secrets in the
  # login keychain are intentionally not removed.
  zap trash: [
    "~/Library/Application Support/SwissKnifeGraph",
    "~/Library/Caches/com.wails.SwissKnifeGraph",
    "~/Library/Preferences/com.wails.SwissKnifeGraph.plist",
    "~/Library/Saved Application State/com.wails.SwissKnifeGraph.savedState",
    "~/Library/WebKit/com.wails.SwissKnifeGraph",
  ]

  caveats <<~EOS
    SwissKnife for MS Graph is not yet signed with an Apple Developer ID or
    notarized, so macOS Gatekeeper may refuse to open it. If that happens run:

      xattr -dr com.apple.quarantine "#{appdir}/SwissKnifeGraph.app"

    Verify the download against SHA256SUMS.txt (minisign-signed) from the
    GitHub release if you want to be sure what you are running.
  EOS
end
