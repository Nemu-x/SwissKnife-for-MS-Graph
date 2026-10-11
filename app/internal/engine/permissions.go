package engine

import "strings"

// Permission preflight: the Graph permissions a manifest names are compared
// with what the connection's token carries, so a tile can warn before a 403
// instead of after it. A broader grant covers a narrower one; anything the
// check cannot reason about (Exchange roles, text hints) is ignored rather
// than reported as missing.

// covers lists grants that imply others. X.ReadWrite.All always implies
// X.Read.All; these are the cross-resource ones.
var covers = map[string][]string{
	"Directory.ReadWrite.All": {"User.ReadWrite.All", "User.Read.All", "Group.ReadWrite.All", "Group.Read.All",
		"GroupMember.ReadWrite.All", "GroupMember.Read.All", "Directory.Read.All"},
	"Directory.Read.All":  {"User.Read.All", "Group.Read.All", "GroupMember.Read.All"},
	"Group.ReadWrite.All": {"GroupMember.ReadWrite.All", "GroupMember.Read.All", "Group.Read.All"},
	"User.ReadWrite.All":  {"User.Read.All"},
	"Mail.ReadWrite":      {"Mail.Read", "Mail.ReadBasic.All", "Mail.ReadBasic"},
	"Mail.Read":           {"Mail.ReadBasic.All", "Mail.ReadBasic"},
	"Calendars.ReadWrite": {"Calendars.Read"},
}

// isGraphPermission tells a plain Graph permission name ("User.ReadWrite.All")
// from descriptive entries ("Exchange.ManageAsApp + Recipient Management").
func isGraphPermission(p string) bool {
	if strings.ContainsAny(p, " +/") || !strings.Contains(p, ".") {
		return false
	}
	return !strings.HasPrefix(p, "Exchange.")
}

func granted(have map[string]bool, want string) bool {
	if have[want] {
		return true
	}
	if strings.HasSuffix(want, ".Read.All") && have[strings.TrimSuffix(want, ".Read.All")+".ReadWrite.All"] {
		return true
	}
	for g, implied := range covers {
		if !have[g] {
			continue
		}
		for _, i := range implied {
			if i == want {
				return true
			}
		}
	}
	return false
}

// missingPermissions returns the manifest's Graph permissions the token
// lacks; have == nil means the grants are unknown and nothing is reported.
func missingPermissions(m Manifest, have map[string]bool) []string {
	if have == nil {
		return nil
	}
	var out []string
	for _, p := range m.Permissions {
		if isGraphPermission(p) && !granted(have, p) {
			out = append(out, p)
		}
	}
	return out
}
