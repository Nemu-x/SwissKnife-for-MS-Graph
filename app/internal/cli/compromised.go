package cli

import (
	"flag"
	"fmt"

	"swissknife-app/internal/services"
)

// setupCompromised: the compromised-account response from a terminal. Block
// and revoke always run; the rest is opt-in, like offboard's actions.
func setupCompromised(fs *flag.FlagSet) func(*env, *globals, []string) int {
	var req services.CompromisedRequest
	fs.StringVar(&req.Confirm, "confirm", "", "repeat the UPN to confirm (required)")
	fs.BoolVar(&req.ResetPassword, "reset-password", false, "set a new temporary password (printed once)")
	fs.BoolVar(&req.ResetMfa, "reset-mfa", false, "remove the registered MFA methods")
	fs.BoolVar(&req.ClearForwarding, "clear-forwarding", false, "stop mailbox forwarding (Exchange PowerShell)")
	fs.BoolVar(&req.DisableRules, "disable-rules", false, "turn off suspicious inbox rules")
	return func(e *env, g *globals, pos []string) int {
		if len(pos) != 1 {
			return report(e, usagef("compromised: expected exactly one UPN"))
		}
		req.Upn = pos[0]
		if req.Confirm != req.Upn {
			return report(e, usagef("compromised: --confirm must repeat the UPN exactly"))
		}
		sess, err := e.session(g)
		if err != nil {
			return report(e, err)
		}
		res, err := services.NewPlaybookService(sess).Compromised(req)
		if err != nil {
			return report(e, err)
		}
		if g.json {
			if code := writeJSON(e, g, res); code != exitOK {
				return code
			}
		} else {
			for _, st := range res.Steps {
				mark := "ok  "
				if !st.OK {
					mark = "FAIL"
				}
				line := fmt.Sprintf("[%s] %s", mark, st.Name)
				if st.Error != "" {
					line += ": " + st.Error
				}
				_, _ = fmt.Fprintln(e.stdout, line)
			}
			if res.TempPassword != "" {
				_, _ = fmt.Fprintf(e.stdout, "Temporary password (shown once, change at next sign-in): %s\n", res.TempPassword)
			}
			for _, s := range res.SignIns {
				_, _ = fmt.Fprintf(e.stdout, "  sign-in %s  %-20s %-15s %s  %s\n", s.When, s.App, s.IP, s.Location, s.Result)
			}
		}
		if !res.OK || res.Canceled {
			return exitFail
		}
		return exitOK
	}
}
