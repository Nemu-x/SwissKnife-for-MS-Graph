// Package cli is the headless entry point of the SwissKnifeGraph binary: the
// same Go services the GUI binds, driven from a terminal or a script. It is
// selected by main() whenever the process receives command-line arguments.
//
// Output is English only (the CLI is not localized). Exit codes: 0 ok,
// 1 command failed, 2 usage error.
package cli

import (
	"context"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"net/url"
	"os"
	"os/signal"
	"path/filepath"
	"strconv"
	"strings"
	"text/tabwriter"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/services"
	"swissknife-app/internal/session"
)

const (
	exitOK    = 0
	exitFail  = 1
	exitUsage = 2
)

const binName = "SwissKnifeGraph"

// env is everything a command needs from the outside world. Run wires the
// real profile store and OS keychain; tests substitute fakes so the command
// layer can be exercised against an httptest Graph without a keychain.
type env struct {
	stdout  io.Writer
	stderr  io.Writer
	version string
	ctx     context.Context
	// profiles lists the saved connection profiles.
	profiles func() ([]secrets.Profile, error)
	// connect returns a session connected with the given profile.
	connect func(ctx context.Context, p secrets.Profile, readOnly bool, stderr io.Writer) (*session.Session, error)
}

// globals are accepted before or after the command name.
type globals struct {
	profile  string
	json     bool
	readOnly bool
}

// bind registers the global flags on fs. The current values serve as defaults
// so a second bind (the command's own FlagSet, after the pre-command parse)
// keeps what was already parsed: flag.StringVar resets the target on registration.
func (g *globals) bind(fs *flag.FlagSet) {
	fs.StringVar(&g.profile, "profile", g.profile, "saved profile id or name (optional when exactly one profile exists)")
	fs.BoolVar(&g.json, "json", g.json, "machine-readable JSON output (default: human-readable text)")
	fs.BoolVar(&g.readOnly, "read-only", g.readOnly, "block every write, like the GUI read-only toggle")
}

// usageError marks a caller mistake (exit 2) as opposed to a failed command (exit 1).
type usageError struct{ err error }

func (u usageError) Error() string { return u.err.Error() }
func (u usageError) Unwrap() error { return u.err }

func usagef(format string, a ...any) error { return usageError{fmt.Errorf(format, a...)} }

// Run executes the CLI with the given arguments (program name excluded) and
// returns the process exit code.
func Run(args []string, version string) int {
	attachConsole()
	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt)
	defer stop()
	e := &env{
		stdout:  os.Stdout,
		stderr:  os.Stderr,
		version: version,
		ctx:     ctx,
		profiles: func() ([]secrets.Profile, error) {
			store, err := secrets.NewStore()
			if err != nil {
				return nil, err
			}
			return store.List()
		},
		connect: connectProfile,
	}
	return run(e, args)
}

// connectProfile mirrors main(): config store, audit log, run journal and
// session, then the same token source the GUI ConnectService builds. Runs
// started here therefore show up in the GUI's audit log and journal too.
func connectProfile(ctx context.Context, p secrets.Profile, readOnly bool, stderr io.Writer) (*session.Session, error) {
	store, err := secrets.NewStore()
	if err != nil {
		return nil, err
	}
	audit := auditlog.New(store.Dir())
	sess := session.New(audit)
	runs := journal.New(filepath.Join(store.Dir(), "runs"))
	runs.Prune(200)
	sess.SetJournal(runs)
	sess.SetConfigDir(store.Dir())
	sess.SetAppContext(ctx)
	sess.SetReadOnly(readOnly)

	cr, err := services.ResolveCredentials(store, services.ConnectRequest{ProfileID: p.ID})
	if err != nil {
		return nil, err
	}
	provider, err := services.NewTokenProvider(cr, func(verifyURL, code, msg string) {
		if msg == "" {
			msg = "To sign in, open " + verifyURL + " and enter the code " + code
		}
		_, _ = fmt.Fprintln(stderr, msg)
	})
	if err != nil {
		return nil, err
	}
	sess.SetClient(graphapi.New(provider), cr.Name)
	sess.Record("session.connect", cr.TenantID, "mode="+cr.AuthMode+" via=cli", nil)
	return sess, nil
}

// command is one sub-command. setup declares its flags on fs (which already
// carries the globals) and returns the executor, so flag variables stay local.
type command struct {
	name    string
	args    string
	summary string
	setup   func(fs *flag.FlagSet) func(e *env, g *globals, pos []string) int
}

// commands is filled in init: "help" walks the table, which would otherwise be
// an initialization cycle.
var commands []command

func init() {
	commands = []command{
		{name: "help", args: "[command]", summary: "show this help or one command's flags", setup: setupHelp},
		{name: "version", summary: "print the application version", setup: setupVersion},
		{name: "profiles", summary: "list saved connection profiles (id, name, tenant, mode)", setup: setupProfiles},
		{name: "get", args: "<graph path> [--top N] [--all]", summary: "raw Graph GET; prints the JSON response (--all follows @odata.nextLink)", setup: setupGet},
		{name: "user", args: "<upn>", summary: "user snapshot: profile, group membership, licenses", setup: setupUser},
		{name: "signins", args: "<upn> [--days 7] [--failed] [--top 50]", summary: "sign-in log for one user", setup: setupSignins},
		{name: "offboard", args: "<upn> --confirm <upn> [actions...]", summary: "run the offboarding playbook (destructive; --confirm must repeat the UPN)", setup: setupOffboard},
	}
}

func findCommand(name string) *command {
	for i := range commands {
		if commands[i].name == name {
			return &commands[i]
		}
	}
	return nil
}

// run is the testable core of Run: global flags, command dispatch, exit code.
func run(e *env, args []string) int {
	var g globals
	pre := flag.NewFlagSet(binName, flag.ContinueOnError)
	pre.SetOutput(io.Discard)
	g.bind(pre)
	if err := pre.Parse(args); err != nil {
		if errors.Is(err, flag.ErrHelp) {
			printUsage(e.stdout)
			return exitOK
		}
		return report(e, usageError{err})
	}
	rest := pre.Args()
	if len(rest) == 0 {
		printUsage(e.stderr)
		return exitUsage
	}
	cmd := findCommand(rest[0])
	if cmd == nil {
		_, _ = fmt.Fprintf(e.stderr, "%s: unknown command %q\n\n", binName, rest[0])
		printUsage(e.stderr)
		return exitUsage
	}

	fs := flag.NewFlagSet(cmd.name, flag.ContinueOnError)
	fs.SetOutput(io.Discard)
	g.bind(fs)
	exec := cmd.setup(fs)
	pos, err := parseInterleaved(fs, rest[1:])
	if errors.Is(err, flag.ErrHelp) {
		printCommandHelp(e.stdout, cmd, fs)
		return exitOK
	}
	if err != nil {
		return report(e, usagef("%s: %v", cmd.name, err))
	}
	return exec(e, &g, pos)
}

// parseInterleaved lets flags and positionals mix ("get /me --top 5"), which
// the stdlib parser does not do on its own: it stops at the first positional.
func parseInterleaved(fs *flag.FlagSet, args []string) ([]string, error) {
	var pos []string
	for {
		if err := fs.Parse(args); err != nil {
			return nil, err
		}
		rest := fs.Args()
		if len(rest) == 0 {
			return pos, nil
		}
		pos = append(pos, rest[0])
		args = rest[1:]
	}
}

// report prints an error and maps it to the exit code.
func report(e *env, err error) int {
	_, _ = fmt.Fprintf(e.stderr, "%s: %v\n", binName, err)
	var ue usageError
	if errors.As(err, &ue) {
		_, _ = fmt.Fprintf(e.stderr, "Run '%s help' for usage.\n", binName)
		return exitUsage
	}
	return exitFail
}

func printUsage(w io.Writer) {
	_, _ = fmt.Fprintf(w, "Usage: %s [global flags] <command> [args] [flags]\n\n", binName)
	_, _ = fmt.Fprintln(w, "Headless mode of SwissKnife for MS Graph: the same Go services as the GUI, driven from a terminal or a script.")
	_, _ = fmt.Fprintln(w, "Tenant commands connect with a saved profile (create it in the app's Connect tab first).")
	_, _ = fmt.Fprintln(w)
	_, _ = fmt.Fprintln(w, "Commands:")
	tw := tabwriter.NewWriter(w, 0, 4, 2, ' ', 0)
	for _, c := range commands {
		syn := c.name
		if c.args != "" {
			syn += " " + c.args
		}
		_, _ = fmt.Fprintf(tw, "  %s\t%s\n", syn, c.summary)
	}
	_ = tw.Flush()
	_, _ = fmt.Fprintln(w)
	_, _ = fmt.Fprintln(w, "Global flags:")
	_, _ = fmt.Fprintln(w, "  --profile <id|name>  saved profile (optional when exactly one profile exists)")
	_, _ = fmt.Fprintln(w, "  --json               machine-readable JSON output")
	_, _ = fmt.Fprintln(w, "  --read-only          block every write")
	_, _ = fmt.Fprintln(w)
	_, _ = fmt.Fprintln(w, "Exit codes: 0 ok, 1 command failed, 2 usage error.")
}

func printCommandHelp(w io.Writer, cmd *command, fs *flag.FlagSet) {
	syn := cmd.name
	if cmd.args != "" {
		syn += " " + cmd.args
	}
	_, _ = fmt.Fprintf(w, "Usage: %s %s\n\n%s\n\nFlags:\n", binName, syn, cmd.summary)
	fs.SetOutput(w)
	fs.PrintDefaults()
	fs.SetOutput(io.Discard)
}

// ---- output helpers ----

// writeJSON prints v: compact for --json (one line, script-friendly), indented otherwise.
func writeJSON(e *env, g *globals, v any) int {
	var (
		b   []byte
		err error
	)
	if g.json {
		b, err = json.Marshal(v)
	} else {
		b, err = json.MarshalIndent(v, "", "  ")
	}
	if err != nil {
		return report(e, err)
	}
	_, _ = fmt.Fprintln(e.stdout, string(b))
	return exitOK
}

func str(m map[string]any, key string) string {
	if v, ok := m[key]; ok && v != nil {
		return fmt.Sprint(v)
	}
	return ""
}

// ---- session ----

// pickProfile resolves --profile by id or (case-insensitive) name; with no
// selector a single saved profile is used implicitly.
func pickProfile(list []secrets.Profile, sel string) (secrets.Profile, error) {
	names := make([]string, 0, len(list))
	for _, p := range list {
		names = append(names, p.Name)
	}
	if sel == "" {
		switch len(list) {
		case 0:
			return secrets.Profile{}, usagef("no saved profiles — create one in the app (Connect tab) first")
		case 1:
			return list[0], nil
		default:
			return secrets.Profile{}, usagef("several profiles exist — pass --profile <id|name> (available: %s)", strings.Join(names, ", "))
		}
	}
	for _, p := range list {
		if p.ID == sel {
			return p, nil
		}
	}
	for _, p := range list {
		if strings.EqualFold(p.Name, sel) {
			return p, nil
		}
	}
	return secrets.Profile{}, usagef("profile %q not found (available: %s)", sel, strings.Join(names, ", "))
}

func (e *env) session(g *globals) (*session.Session, error) {
	list, err := e.profiles()
	if err != nil {
		return nil, err
	}
	p, err := pickProfile(list, g.profile)
	if err != nil {
		return nil, err
	}
	return e.connect(e.ctx, p, g.readOnly, e.stderr)
}

// ---- commands ----

func setupHelp(fs *flag.FlagSet) func(*env, *globals, []string) int {
	return func(e *env, g *globals, pos []string) int {
		if len(pos) == 0 {
			printUsage(e.stdout)
			return exitOK
		}
		cmd := findCommand(pos[0])
		if cmd == nil {
			return report(e, usagef("unknown command %q", pos[0]))
		}
		cfs := flag.NewFlagSet(cmd.name, flag.ContinueOnError)
		g.bind(cfs)
		cmd.setup(cfs)
		printCommandHelp(e.stdout, cmd, cfs)
		return exitOK
	}
}

func setupVersion(fs *flag.FlagSet) func(*env, *globals, []string) int {
	return func(e *env, g *globals, pos []string) int {
		if g.json {
			return writeJSON(e, g, map[string]string{"name": binName, "version": e.version})
		}
		_, _ = fmt.Fprintf(e.stdout, "%s %s\n", binName, e.version)
		return exitOK
	}
}

func setupProfiles(fs *flag.FlagSet) func(*env, *globals, []string) int {
	return func(e *env, g *globals, pos []string) int {
		list, err := e.profiles()
		if err != nil {
			return report(e, err)
		}
		if list == nil {
			list = []secrets.Profile{}
		}
		if g.json {
			return writeJSON(e, g, list)
		}
		if len(list) == 0 {
			_, _ = fmt.Fprintln(e.stdout, "No saved profiles. Create one in the app (Connect tab, \"Remember this profile\").")
			return exitOK
		}
		tw := tabwriter.NewWriter(e.stdout, 0, 4, 2, ' ', 0)
		_, _ = fmt.Fprintln(tw, "ID\tNAME\tTENANT\tMODE")
		for _, p := range list {
			_, _ = fmt.Fprintf(tw, "%s\t%s\t%s\t%s\n", p.ID, p.Name, p.TenantID, p.AuthMode)
		}
		_ = tw.Flush()
		return exitOK
	}
}

func setupGet(fs *flag.FlagSet) func(*env, *globals, []string) int {
	top := fs.Int("top", 0, "page size ($top)")
	all := fs.Bool("all", false, "follow @odata.nextLink and print every item as {\"value\": [...]}")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 1 {
			return report(e, usagef("get: expected exactly one Graph path, e.g. /users?$select=id,displayName"))
		}
		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}
		c, err := sess.Client()
		if err != nil {
			return report(e, err)
		}
		var params url.Values
		if *top > 0 {
			params = url.Values{"$top": {strconv.Itoa(*top)}}
		}
		if *all {
			items, err := c.ListAll(sess.Ctx(), pos[0], params, 0)
			if err != nil {
				return report(e, err)
			}
			if items == nil {
				items = []json.RawMessage{}
			}
			return writeJSON(e, g, map[string]any{"value": items})
		}
		var out json.RawMessage
		if err := c.Get(sess.Ctx(), pos[0], params, &out); err != nil {
			return report(e, err)
		}
		if out == nil {
			out = json.RawMessage(`{"status":"no content"}`)
		}
		return writeJSON(e, g, out)
	}
}

func setupUser(fs *flag.FlagSet) func(*env, *globals, []string) int {
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 1 {
			return report(e, usagef("user: expected exactly one UPN or object id"))
		}
		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}
		snap, err := services.NewUsersService(sess).Snapshot(pos[0])
		if err != nil {
			return report(e, err)
		}
		if g.json {
			return writeJSON(e, g, snap)
		}
		printSnapshot(e.stdout, snap)
		return exitOK
	}
}

// printSnapshot renders the Snapshot map (profile RawMessage + memberOf and
// licenses RawMessage lists) as a short human summary.
func printSnapshot(w io.Writer, snap map[string]any) {
	var profile map[string]any
	if raw, ok := snap["profile"].(json.RawMessage); ok {
		_ = json.Unmarshal(raw, &profile)
	}
	tw := tabwriter.NewWriter(w, 0, 4, 2, ' ', 0)
	_, _ = fmt.Fprintf(tw, "User:\t%s (%s)\n", str(profile, "displayName"), str(profile, "userPrincipalName"))
	_, _ = fmt.Fprintf(tw, "Id:\t%s\n", str(profile, "id"))
	_, _ = fmt.Fprintf(tw, "Enabled:\t%s\n", str(profile, "accountEnabled"))
	_, _ = fmt.Fprintf(tw, "Mail:\t%s\n", str(profile, "mail"))
	_, _ = fmt.Fprintf(tw, "Job title:\t%s\n", str(profile, "jobTitle"))
	_, _ = fmt.Fprintf(tw, "Department:\t%s\n", str(profile, "department"))
	_ = tw.Flush()

	groups := decodeList(snap["memberOf"])
	_, _ = fmt.Fprintf(w, "Groups (%d):\n", len(groups))
	for _, m := range groups {
		kind := strings.TrimPrefix(str(m, "@odata.type"), "#microsoft.graph.")
		_, _ = fmt.Fprintf(w, "  - %s (%s)\n", str(m, "displayName"), kind)
	}
	licenses := decodeList(snap["licenses"])
	_, _ = fmt.Fprintf(w, "Licenses (%d):\n", len(licenses))
	for _, l := range licenses {
		_, _ = fmt.Fprintf(w, "  - %s\n", str(l, "skuPartNumber"))
	}
}

func decodeList(v any) []map[string]any {
	raws, ok := v.([]json.RawMessage)
	if !ok {
		return nil
	}
	out := make([]map[string]any, 0, len(raws))
	for _, r := range raws {
		var m map[string]any
		if json.Unmarshal(r, &m) == nil {
			out = append(out, m)
		}
	}
	return out
}

func setupSignins(fs *flag.FlagSet) func(*env, *globals, []string) int {
	days := fs.Int("days", 7, "look back this many days (0 = no limit)")
	failed := fs.Bool("failed", false, "only sign-ins Entra rejected")
	top := fs.Int("top", 50, "maximum number of events")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 1 {
			return report(e, usagef("signins: expected exactly one UPN"))
		}
		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}
		items, err := services.NewAuditService(sess).SignInsFiltered(services.SignInQuery{
			Upn: pos[0], Days: *days, FailedOnly: *failed, Top: *top,
		})
		if err != nil {
			return report(e, err)
		}
		if items == nil {
			items = []json.RawMessage{}
		}
		if g.json {
			return writeJSON(e, g, items)
		}
		type signIn struct {
			CreatedDateTime string `json:"createdDateTime"`
			AppDisplayName  string `json:"appDisplayName"`
			IPAddress       string `json:"ipAddress"`
			Status          struct {
				ErrorCode     int    `json:"errorCode"`
				FailureReason string `json:"failureReason"`
			} `json:"status"`
			Location struct {
				City            string `json:"city"`
				CountryOrRegion string `json:"countryOrRegion"`
			} `json:"location"`
		}
		tw := tabwriter.NewWriter(e.stdout, 0, 4, 2, ' ', 0)
		_, _ = fmt.Fprintln(tw, "TIME\tRESULT\tAPP\tIP\tLOCATION\tREASON")
		for _, raw := range items {
			var s signIn
			if err := json.Unmarshal(raw, &s); err != nil {
				continue
			}
			result := "ok"
			if s.Status.ErrorCode != 0 {
				result = "FAIL " + strconv.Itoa(s.Status.ErrorCode)
			}
			loc := strings.Trim(s.Location.City+", "+s.Location.CountryOrRegion, ", ")
			_, _ = fmt.Fprintf(tw, "%s\t%s\t%s\t%s\t%s\t%s\n", s.CreatedDateTime, result, s.AppDisplayName, s.IPAddress, loc, s.Status.FailureReason)
		}
		_ = tw.Flush()
		_, _ = fmt.Fprintf(e.stdout, "%d event(s)\n", len(items))
		return exitOK
	}
}

func setupOffboard(fs *flag.FlagSet) func(*env, *globals, []string) int {
	var req services.OffboardRequest
	fs.StringVar(&req.Confirm, "confirm", "", "repeat the UPN to confirm (required)")
	fs.BoolVar(&req.Block, "block", false, "block sign-in")
	fs.BoolVar(&req.RevokeSessions, "revoke", false, "revoke refresh tokens / sessions")
	fs.StringVar(&req.OofMessage, "oof", "", "set an automatic reply with this text")
	fs.StringVar(&req.ForwardTo, "forward", "", "forward incoming mail to this UPN (inbox rule)")
	fs.BoolVar(&req.HideFromGal, "hide-gal", false, "hide from address lists")
	fs.StringVar(&req.CalendarTo, "share-calendar", "", "share the calendar (read) with this UPN")
	fs.BoolVar(&req.RemoveFromGroups, "remove-groups", false, "remove from all groups")
	fs.BoolVar(&req.RemoveAllLicenses, "remove-licenses", false, "remove all licenses")
	fs.StringVar(&req.BackupToUser, "backup-to", "", "copy OneDrive to this user's drive (server-side)")
	fs.StringVar(&req.BackupFolder, "backup-folder", "", "target folder for the backup (default: the leaver's UPN)")
	fs.BoolVar(&req.BackupChats, "backup-chats", false, "also export Teams chats into the backup (needs --backup-to)")
	fs.StringVar(&req.IntuneAction, "intune", "", "Intune action for the user's devices: retire | wipe")
	fs.BoolVar(&req.RemoveMfaMethods, "remove-mfa", false, "remove registered MFA methods")
	fs.BoolVar(&req.DeleteRegisteredDevices, "delete-devices", false, "delete the user's registered Entra devices")
	fs.StringVar(&req.TransferOwnershipTo, "transfer-ownership", "", "make this UPN owner of the leaver's groups and teams")
	fs.BoolVar(&req.CancelFutureEvents, "cancel-events", false, "cancel meetings the leaver organised from now on")
	fs.BoolVar(&req.Delete, "delete", false, "delete the user account (last step)")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 1 {
			return report(e, usagef("offboard: expected exactly one UPN"))
		}
		req.Upn = pos[0]
		if req.Confirm == "" {
			return report(e, usagef("offboard: --confirm <upn> is required"))
		}
		if req.Confirm != req.Upn {
			return report(e, usagef("offboard: --confirm must repeat the UPN exactly"))
		}
		switch req.IntuneAction {
		case "", "retire", "wipe":
		default:
			return report(e, usagef("offboard: --intune must be retire or wipe"))
		}
		req.Oof = req.OofMessage != ""
		if req.BackupChats && req.BackupToUser == "" {
			return report(e, usagef("offboard: --backup-chats needs --backup-to"))
		}
		hasAction := req.Block || req.RevokeSessions || req.Oof || req.ForwardTo != "" || req.HideFromGal ||
			req.CalendarTo != "" || req.RemoveFromGroups || req.RemoveAllLicenses || req.BackupToUser != "" ||
			req.IntuneAction != "" || req.RemoveMfaMethods || req.DeleteRegisteredDevices ||
			req.TransferOwnershipTo != "" || req.CancelFutureEvents || req.Delete
		if !hasAction {
			return report(e, usagef("offboard: nothing to do — pass at least one action flag (see '%s help offboard')", binName))
		}

		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}

		// Stream completed steps to stderr as they happen; the summary/JSON
		// result goes to stdout once the run is over.
		services.SetEventSink(func(name string, data map[string]any) {
			if name != "playbook:step" || data["status"] != "done" {
				return
			}
			mark := "ok  "
			if ok, _ := data["ok"].(bool); !ok {
				mark = "FAIL"
			}
			line := fmt.Sprintf("[%s] %s", mark, str(data, "name"))
			if d := str(data, "detail"); d != "" {
				line += " — " + d
			}
			if er := str(data, "error"); er != "" {
				line += ": " + er
			}
			if h := str(data, "hint"); h != "" {
				line += " (" + h + ")"
			}
			_, _ = fmt.Fprintln(e.stderr, line)
		})
		defer services.SetEventSink(nil)

		res, err := services.NewPlaybookService(sess).Offboard(req)
		if err != nil {
			return report(e, err)
		}
		if g.json {
			code := writeJSON(e, g, res)
			if code != exitOK || !res.OK || res.Canceled {
				return exitFail
			}
			return exitOK
		}
		failed := 0
		for _, st := range res.Steps {
			if !st.OK {
				failed++
			}
		}
		switch {
		case res.Canceled:
			_, _ = fmt.Fprintf(e.stdout, "Offboarding %s canceled after %d step(s).\n", req.Upn, len(res.Steps))
			return exitFail
		case failed > 0:
			_, _ = fmt.Fprintf(e.stdout, "Offboarding %s finished with %d of %d step(s) failed.\n", req.Upn, failed, len(res.Steps))
			return exitFail
		default:
			_, _ = fmt.Fprintf(e.stdout, "Offboarding %s completed: %d step(s) ok.\n", req.Upn, len(res.Steps))
			return exitOK
		}
	}
}
