package cli

import (
	"flag"
	"fmt"
	"strings"
	"text/tabwriter"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/services"
)

// setupAction drives the action catalog (ADR-008) from a terminal:
//
//	action list
//	action <id> field=value ... [--apply] [--confirm <target>]
//
// Without --apply only the preview is printed, so a script can show what would
// change before committing to it — the same two steps as the GUI.
func setupAction(fs *flag.FlagSet) func(*env, *globals, []string) int {
	apply := fs.Bool("apply", false, "apply the previewed changes (default: preview only)")
	confirm := fs.String("confirm", "", "repeat the target to apply a destructive action")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) == 0 {
			return report(e, usagef("action: expected 'list' or an action id"))
		}
		in := engine.Inputs{}
		for _, kv := range pos[1:] {
			k, v, ok := strings.Cut(kv, "=")
			if !ok || k == "" {
				return report(e, usagef("action: expected field=value, got %q", kv))
			}
			in[k] = v
		}

		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}
		eng := services.NewEngine(sess)

		if pos[0] == "list" {
			return printCatalog(e, g, eng.Catalog())
		}

		if isRead(eng.Catalog(), pos[0]) {
			res, err := eng.Run(e.ctx, pos[0], in)
			if err != nil {
				return report(e, err)
			}
			if g.json {
				return writeJSON(e, g, res)
			}
			printRows(e, res)
			return exitOK
		}

		plan, err := eng.Plan(pos[0], in)
		if err != nil {
			return report(e, err)
		}
		if !*apply {
			if g.json {
				return writeJSON(e, g, map[string]any{"plan": plan})
			}
			printPlan(e, plan)
			_, _ = fmt.Fprintln(e.stdout, "Preview only — run again with --apply to make these changes.")
			return exitOK
		}

		res, err := eng.Apply(plan.ID, *confirm)
		if err != nil {
			return report(e, err)
		}
		if g.json {
			if code := writeJSON(e, g, map[string]any{"plan": plan, "result": res}); code != exitOK {
				return code
			}
		} else {
			printPlan(e, plan)
			_, _ = fmt.Fprintf(e.stdout, "Applied %d, unchanged %d, failed %d.\n", res.Applied, res.Skipped, res.Failed)
			for _, o := range res.Outcomes {
				if !o.OK {
					_, _ = fmt.Fprintf(e.stderr, "FAIL %s %s: %s\n", o.Target, o.Field, o.Error)
				}
			}
		}
		if res.Failed > 0 || res.Canceled {
			return exitFail
		}
		return exitOK
	}
}

func printCatalog(e *env, g *globals, list []engine.CatalogEntry) int {
	if g.json {
		return writeJSON(e, g, list)
	}
	tw := tabwriter.NewWriter(e.stdout, 0, 4, 2, ' ', 0)
	_, _ = fmt.Fprintln(tw, "ID\tDANGER\tFIELDS\tSTATUS")
	for _, a := range list {
		fields := make([]string, 0, len(a.Fields))
		for _, f := range a.Fields {
			name := f.Name
			if len(f.Options) > 0 {
				name += "=" + strings.Join(f.Options, "|")
			}
			if !f.Required {
				name = "[" + name + "]"
			}
			fields = append(fields, name)
		}
		status := "ok (" + string(a.Backend) + ")"
		if !a.Available && a.Reason != nil {
			status = "unavailable: " + a.Reason.Key
		}
		_, _ = fmt.Fprintf(tw, "%s\t%s\t%s\t%s\n", a.ID, a.Danger, strings.Join(fields, " "), status)
	}
	_ = tw.Flush()
	return exitOK
}

func isRead(list []engine.CatalogEntry, id string) bool {
	for _, a := range list {
		if a.ID == id {
			return a.Danger == engine.Read
		}
	}
	return false
}

func printRows(e *env, r *engine.ReadResult) {
	tw := tabwriter.NewWriter(e.stdout, 0, 4, 2, ' ', 0)
	_, _ = fmt.Fprintln(tw, strings.ToUpper(strings.Join(r.Columns, "	")))
	for _, row := range r.Rows {
		vals := make([]string, len(r.Columns))
		for i, c := range r.Columns {
			vals[i] = row[c]
		}
		_, _ = fmt.Fprintln(tw, strings.Join(vals, "	"))
	}
	_ = tw.Flush()
}

func printPlan(e *env, p *engine.Plan) {
	_, _ = fmt.Fprintf(e.stdout, "Plan for %s (%s):\n", p.ActionID, p.Backend)
	for _, c := range p.Changes {
		mark, what := "~", fmt.Sprintf("%s → %s", dash(c.Before), dash(c.After))
		switch c.Op {
		case "add":
			mark, what = "+", dash(c.After)
		case "remove":
			mark, what = "-", dash(c.Before)
		case "none":
			mark, what = "=", dash(firstNonEmpty(c.After, c.Before))+" (already so)"
		}
		line := fmt.Sprintf("  %s %s  %s: %s", mark, c.Target, c.Field, what)
		if c.Note != "" {
			line += " [" + c.Note + "]"
		}
		_, _ = fmt.Fprintln(e.stdout, line)
	}
}

func dash(s string) string {
	if s == "" {
		return "—"
	}
	return s
}

func firstNonEmpty(a, b string) string {
	if a != "" {
		return a
	}
	return b
}
