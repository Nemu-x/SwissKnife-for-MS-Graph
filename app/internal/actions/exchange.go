package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/exoapi"
)

// Exchange actions. The Admin API implementation comes first; the
// PowerShell implementations join these lists with the pwsh backend.

func exchangeActions() []engine.Action {
	return []engine.Action{
		{
			Manifest: engine.Manifest{
				ID: "mailbox.sendOnBehalf", Page: "mail", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "mailbox", Kind: engine.FieldUser, Required: true},
					{Name: "delegate", Kind: engine.FieldUser, Required: true},
					{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
				},
				Permissions: []string{"Exchange.ManageAsAppV2 + Recipient Management"},
			},
			Impls: []engine.Impl{exoSendOnBehalf{}, psSendOnBehalf{}},
		},
		{
			Manifest: engine.Manifest{
				ID: "mailbox.folderPermission", Page: "mail", Danger: engine.Write,
				Fields: []engine.Field{
					{Name: "mailbox", Kind: engine.FieldUser, Required: true},
					{Name: "folder", Kind: engine.FieldChoice, Required: true, Options: []string{"calendar", "inbox"}, Default: "calendar"},
					{Name: "user", Kind: engine.FieldUser, Required: true},
					{Name: "access", Kind: engine.FieldChoice, Required: true,
						Options: []string{"AvailabilityOnly", "LimitedDetails", "Reviewer", "Author", "Editor", "Owner", "none"}, Default: "Reviewer"},
				},
				Permissions: []string{"Exchange.ManageAsAppV2 + Recipient Management", "Calendars.Read", "Mail.ReadBasic.All"},
			},
			Impls: []engine.Impl{exoFolderPermission{}, psFolderPermission{}},
		},
	}
}

// recipient is a Graph user reduced to the identifiers Exchange may echo
// back: GrantSendOnBehalfTo lists recipient names, which are the alias or,
// for newer objects, the Entra object id.
type recipient struct {
	ID, UPN, Mail, Name, Alias string
}

// is matches any identifier, display name included (weakest, last resort).
func (r recipient) is(v string) bool {
	return r.isAddress(v) || (v != "" && strings.EqualFold(v, r.Name))
}

// isAddress matches identifiers that are unique in the tenant.
func (r recipient) isAddress(v string) bool {
	if v == "" {
		return false
	}
	for _, x := range []string{r.ID, r.UPN, r.Mail, r.Alias} {
		if x != "" && strings.EqualFold(v, x) {
			return true
		}
	}
	return false
}

func lookupRecipient(env engine.Env, user string) (recipient, error) {
	var u struct {
		ID    string `json:"id"`
		UPN   string `json:"userPrincipalName"`
		Mail  string `json:"mail"`
		Name  string `json:"displayName"`
		Alias string `json:"mailNickname"`
	}
	err := env.Graph.Get(env.Ctx, "/users/"+url.PathEscape(user), url.Values{"$select": {"id,userPrincipalName,mail,displayName,mailNickname"}}, &u)
	if err != nil {
		return recipient{}, err
	}
	if u.UPN == "" {
		u.UPN = user
	}
	return recipient{ID: u.ID, UPN: u.UPN, Mail: u.Mail, Name: u.Name, Alias: u.Alias}, nil
}

// --- mailbox.sendOnBehalf --------------------------------------------------

type exoSendOnBehalf struct{}

func (exoSendOnBehalf) Backend() engine.Backend { return exoapi.BackendExoAPI }

func (exoSendOnBehalf) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	c, err := exoapi.FromEnv(env)
	if err != nil {
		return nil, err
	}
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	delegate, err := lookupRecipient(env, in["delegate"])
	if err != nil {
		return nil, err
	}
	items, err := c.Invoke(env.Ctx, "Mailbox", "Get-Mailbox", map[string]any{"Identity": owner.UPN}, exoapi.MailboxAnchor(owner.UPN))
	if err != nil {
		return nil, err
	}
	if len(items) == 0 {
		return nil, fmt.Errorf("no mailbox found for %s", owner.UPN)
	}
	var mb struct {
		UPN       string   `json:"UserPrincipalName"`
		Delegates []string `json:"GrantSendOnBehalfTo"`
	}
	if err := json.Unmarshal(items[0], &mb); err != nil {
		return nil, err
	}
	has := false
	for _, d := range mb.Delegates {
		if delegate.is(d) {
			has = true
		}
	}
	target := firstOf(mb.UPN, owner.UPN)
	ch := engine.Change{Target: target, Field: "sendOnBehalf", Op: in["op"],
		Ref: map[string]string{"mailbox": owner.UPN, "delegate": firstOf(delegate.Mail, delegate.UPN)}}
	switch {
	case in["op"] == "add" && has:
		ch.Op, ch.Before, ch.After = "none", delegate.UPN, delegate.UPN
	case in["op"] == "add":
		ch.After = delegate.UPN
	case has:
		ch.Before = delegate.UPN
	default:
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (exoSendOnBehalf) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	c, err := exoapi.FromEnv(env)
	if err != nil {
		return err
	}
	delta := map[string]any{"@odata.type": "#Exchange.GenericHashTable", ch.Op: []string{ch.Ref["delegate"]}}
	_, err = c.Invoke(env.Ctx, "Mailbox", "Set-Mailbox",
		map[string]any{"Identity": ch.Ref["mailbox"], "GrantSendOnBehalfTo": delta}, exoapi.MailboxAnchor(ch.Ref["mailbox"]))
	return err
}

// --- mailbox.folderPermission ----------------------------------------------

type exoFolderPermission struct{}

func (exoFolderPermission) Backend() engine.Backend { return exoapi.BackendExoAPI }

// calendarOnly roles exist only on calendar folders.
var calendarOnly = map[string]bool{"AvailabilityOnly": true, "LimitedDetails": true}

// folderIdentity builds "<mailbox>:\<folder>" with the folder's real
// (localized) name: a Russian mailbox has "Календарь", not "Calendar".
//
// Guessing the English name instead would fail later with a confusing
// "folder not found", so a Graph error is returned as is (with its
// permission hint).
func folderIdentity(env engine.Env, mailbox, folder string) (string, error) {
	var f struct {
		Name        string `json:"name"`
		DisplayName string `json:"displayName"`
	}
	path, sel := "/calendar", "name"
	if folder == "inbox" {
		path, sel = "/mailFolders/inbox", "displayName"
	}
	if err := env.Graph.Get(env.Ctx, "/users/"+url.PathEscape(mailbox)+path, url.Values{"$select": {sel}}, &f); err != nil {
		return "", err
	}
	name := firstOf(f.Name, f.DisplayName)
	if name == "" {
		return "", fmt.Errorf("the %s folder of %s has no name", folder, mailbox)
	}
	return mailbox + `:` + "\\" + name, nil
}

// permEntry decodes one Get-MailboxFolderPermission row; User may come as a
// plain string or as an object depending on the service version.
type permEntry struct {
	User         json.RawMessage `json:"User"`
	AccessRights json.RawMessage `json:"AccessRights"` // string or array
}

// names returns the row's address (when the service sent one) and its
// display name. User arrives as a plain string or as an object.
func (p permEntry) names() (address, display string) {
	var s string
	if json.Unmarshal(p.User, &s) == nil {
		if strings.Contains(s, "@") {
			return s, ""
		}
		return "", s
	}
	var o struct {
		DisplayName string `json:"DisplayName"`
		Recipient   struct {
			PrimarySmtpAddress string `json:"PrimarySmtpAddress"`
		} `json:"ADRecipient"`
	}
	_ = json.Unmarshal(p.User, &o)
	return o.Recipient.PrimarySmtpAddress, o.DisplayName
}

// findGrantee picks the grantee's row: an address match wins; a display
// name counts only when exactly one row carries it, because two people can
// share a name and the wrong row would plan the wrong change.
func findGrantee(rows []permEntry, r recipient) (permEntry, bool) {
	var byName []permEntry
	for _, row := range rows {
		addr, display := row.names()
		if r.isAddress(addr) {
			return row, true
		}
		// A row with a different address is someone else, whatever its name.
		if addr == "" && display != "" && (r.isAddress(display) || strings.EqualFold(display, r.Name)) {
			byName = append(byName, row)
		}
	}
	if len(byName) == 1 {
		return byName[0], true
	}
	return permEntry{}, false
}

func (exoFolderPermission) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	if in["folder"] != "calendar" && calendarOnly[in["access"]] {
		return nil, errors.New(in["access"] + " exists only on calendar folders")
	}
	c, err := exoapi.FromEnv(env)
	if err != nil {
		return nil, err
	}
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	grantee, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	identity, err := folderIdentity(env, owner.UPN, in["folder"])
	if err != nil {
		return nil, err
	}
	items, err := c.Invoke(env.Ctx, "MailboxFolderPermission", "Get-MailboxFolderPermission",
		map[string]any{"Identity": identity}, exoapi.MailboxAnchor(owner.UPN))
	if err != nil {
		return nil, err
	}
	return planFolderChange(items, grantee, identity, owner.UPN, in["access"]), nil
}

// planFolderChange turns the folder's permission rows into the change for
// one grantee; shared by the Admin API and PowerShell implementations.
func planFolderChange(items []json.RawMessage, grantee recipient, identity, mailbox, want string) []engine.Change {
	rows := make([]permEntry, 0, len(items))
	for _, raw := range items {
		var e permEntry
		if json.Unmarshal(raw, &e) == nil {
			rows = append(rows, e)
		}
	}
	current := ""
	if row, ok := findGrantee(rows, grantee); ok {
		current = strings.Join(strs(row.AccessRights), ",")
	}
	ch := engine.Change{Target: grantee.UPN + " @ " + identity, Field: "folderAccess", Before: current,
		Ref: map[string]string{"identity": identity, "mailbox": mailbox, "user": firstOf(grantee.Mail, grantee.UPN)}}
	switch {
	case want == "none" && current == "":
		ch.Op = "none"
	case want == "none":
		ch.Op = "remove"
	case current == "":
		ch.Op, ch.After = "add", want
	case strings.EqualFold(current, want):
		ch.Op, ch.After = "none", want
	default:
		ch.Op, ch.After = "set", want
	}
	return []engine.Change{ch}
}

func (exoFolderPermission) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	c, err := exoapi.FromEnv(env)
	if err != nil {
		return err
	}
	params := map[string]any{"Identity": ch.Ref["identity"], "User": ch.Ref["user"]}
	cmdlet := "Remove-MailboxFolderPermission"
	switch ch.Op {
	case "add":
		cmdlet, params["AccessRights"] = "Add-MailboxFolderPermission", ch.After
	case "set":
		cmdlet, params["AccessRights"] = "Set-MailboxFolderPermission", ch.After
	}
	_, err = c.Invoke(env.Ctx, "MailboxFolderPermission", cmdlet, params, exoapi.MailboxAnchor(ch.Ref["mailbox"]))
	return err
}
