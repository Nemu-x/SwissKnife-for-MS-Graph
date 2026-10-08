package cli

import (
	"context"
	"flag"
	"fmt"
	"os"
	"os/signal"
	"path/filepath"
	"strings"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/worker"
)

// setupWorker runs this machine as a PowerShell worker for paired clients:
//
//	worker serve [--listen :8743] [--families exo,teams] [--pair]
//	worker clients
//	worker revoke <fingerprint>
func setupWorker(fs *flag.FlagSet) func(*env, *globals, []string) int {
	listen := fs.String("listen", ":8743", "address to listen on")
	families := fs.String("families", "exo,teams", "module families to run for clients (exo, teams)")
	pair := fs.Bool("pair", false, "open pairing for one client and print its code (10 minutes)")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) == 0 {
			_, _ = fmt.Fprintln(e.stderr, "usage: worker serve [--listen :8743] [--families exo,teams] [--pair] | worker clients | worker revoke <fingerprint>")
			return exitUsage
		}
		base, err := os.UserConfigDir()
		if err != nil {
			_, _ = fmt.Fprintln(e.stderr, err)
			return exitFail
		}
		dir := filepath.Join(base, "SwissKnifeGraph", "worker")
		id, err := worker.LoadOrCreateIdentity(dir, "worker")
		if err != nil {
			_, _ = fmt.Fprintln(e.stderr, err)
			return exitFail
		}
		det := pwsh.NewDetector()
		allow := actions.PowerShellCmdlets()
		fam := map[string]bool{}
		for _, f := range strings.Split(*families, ",") {
			if f = strings.TrimSpace(f); f != "" {
				if _, ok := allow[f]; !ok {
					_, _ = fmt.Fprintf(e.stderr, "unknown family %q (exo, teams)\n", f)
					return exitUsage
				}
				fam[f] = true
			}
		}
		pool := pwsh.NewPool(det, allow)
		defer pool.Close()
		srv := &worker.Server{Identity: id, Dir: dir, Version: e.version, Runner: pool, Families: fam, Allow: allow,
			Audit: auditlog.New(dir), Logf: func(f string, a ...any) { _, _ = fmt.Fprintf(e.stderr, f+"\n", a...) }}
		srv.Name, _ = os.Hostname()
		srv.Available = func() map[string]bool {
			pe := det.Get(e.ctx)
			return map[string]bool{
				pwsh.FamilyExchange: pe.PwshOK() && pe.Supports(pwsh.ModuleExchange),
				pwsh.FamilyTeams:    pe.PwshOK() && pe.Supports(pwsh.ModuleTeams),
			}
		}
		switch pos[0] {
		case "clients":
			list := srv.Clients()
			if g.json {
				return writeJSON(e, g, list)
			}
			for _, c := range list {
				_, _ = fmt.Fprintf(e.stdout, "%s  %s  paired %s\n", c.Fingerprint[:16], c.Name, c.PairedAt.Format("2006-01-02"))
			}
			return exitOK
		case "revoke":
			if len(pos) != 2 {
				_, _ = fmt.Fprintln(e.stderr, "usage: worker revoke <fingerprint>")
				return exitUsage
			}
			if err := srv.Revoke(pos[1]); err != nil {
				_, _ = fmt.Fprintln(e.stderr, err)
				return exitFail
			}
			_, _ = fmt.Fprintln(e.stdout, "revoked")
			return exitOK
		case "serve":
		default:
			_, _ = fmt.Fprintf(e.stderr, "unknown worker command %q\n", pos[0])
			return exitUsage
		}
		_, _ = fmt.Fprintf(e.stdout, "SwissKnife worker %s — certificate %s\n", srv.Name, id.Fingerprint)
		_, _ = fmt.Fprintf(e.stdout, "runs: %v (installed: %v)\n", strings.Split(*families, ","), srv.Available())
		if *pair {
			code := srv.StartPairing()
			_, _ = fmt.Fprintf(e.stdout, "\npairing code: %s  (valid 10 minutes, one client)\n"+
				"In SwissKnife: Settings → Remote worker → address of this machine and the code.\n"+
				"Check that the app then shows this certificate: %s…\n\n", code, id.Fingerprint[:16])
		}
		ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt)
		defer stop()
		_, _ = fmt.Fprintf(e.stdout, "listening on %s — Ctrl+C to stop\n", *listen)
		if err := srv.Serve(ctx, *listen); err != nil {
			_, _ = fmt.Fprintln(e.stderr, err)
			return exitFail
		}
		return exitOK
	}
}
