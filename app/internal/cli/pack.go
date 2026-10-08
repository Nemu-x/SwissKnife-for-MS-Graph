package cli

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"swissknife-app/internal/packs"
)

// setupPack is the pack author's helper: `pack digest <dir>` writes the
// pack.digest file that minisign then signs (minisign -Sm pack.digest).
func setupPack(fs *flag.FlagSet) func(*env, *globals, []string) int {
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 2 || pos[0] != "digest" {
			_, _ = fmt.Fprintln(e.stderr, "usage: pack digest <pack folder>")
			return exitUsage
		}
		dir := pos[1]
		if _, err := os.Stat(filepath.Join(dir, packs.ManifestFile)); err != nil {
			_, _ = fmt.Fprintf(e.stderr, "%s: no %s\n", dir, packs.ManifestFile)
			return exitUsage
		}
		d, err := packs.Digest(dir)
		if err != nil {
			_, _ = fmt.Fprintln(e.stderr, err)
			return exitFail
		}
		if err := os.WriteFile(filepath.Join(dir, packs.DigestFile), []byte(d+"\n"), 0o644); err != nil {
			_, _ = fmt.Fprintln(e.stderr, err)
			return exitFail
		}
		if g.json {
			return writeJSON(e, g, map[string]string{"digest": d, "file": filepath.Join(dir, packs.DigestFile)})
		}
		_, _ = fmt.Fprintf(e.stdout, "%s\nwritten to %s — sign it with: minisign -Sm %s\n", d,
			filepath.Join(dir, packs.DigestFile), filepath.Join(dir, packs.DigestFile))
		return exitOK
	}
}
