package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"regexp"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/pwsh"
)

// Exchange Online PowerShell implementations. They cover everything the Admin
// API does not (FullAccess, Send As, mailbox type, forwarding, addresses,
// calendar processing, distribution lists, transport rules) and stand in for
// the Admin API where a tenant lacks the Preview.

// exchangeCmdlets is the host allow-list for the Exchange family: nothing
// else can run there.
var exchangeCmdlets = []string{
	"Get-Mailbox", "Set-Mailbox",
	"Get-MailboxFolderPermission", "Add-MailboxFolderPermission", "Set-MailboxFolderPermission", "Remove-MailboxFolderPermission",
	"Get-MailboxPermission", "Add-MailboxPermission", "Remove-MailboxPermission",
	"Get-RecipientPermission", "Add-RecipientPermission", "Remove-RecipientPermission",
	"Get-CalendarProcessing", "Set-CalendarProcessing",
	"Get-DistributionGroupMember", "Add-DistributionGroupMember", "Remove-DistributionGroupMember",
	"Get-TransportRule", "Enable-TransportRule", "Disable-TransportRule",
}

// PowerShellCmdlets returns the allow-list per module family for the host pool.
func PowerShellCmdlets() map[string][]string {
	return map[string][]string{
		pwsh.FamilyExchange: append(append([]string{}, exchangeCmdlets...), securityCmdlets...),
		pwsh.FamilyTeams:    teamsCmdlets(),
	}
}

func exo(env engine.Env, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error) {
	if env.PS == nil {
		return nil, errors.New("the PowerShell backend is not set up")
	}
	return env.PS.Invoke(env, pwsh.FamilyExchange, cmdlet, params, sel...)
}

// strs reads a PowerShell value that may be a string, an array of strings, or
// an array of objects with a Name / PrimarySmtpAddress.
func strs(raw json.RawMessage) []string {
	var one string
	if json.Unmarshal(raw, &one) == nil {
		if one == "" {
			return nil
		}
		return []string{one}
	}
	var many []json.RawMessage
	if json.Unmarshal(raw, &many) != nil {
		return nil
	}
	var out []string
	for _, m := range many {
		var s string
		if json.Unmarshal(m, &s) == nil {
			out = append(out, s)
			continue
		}
		var o struct{ Name, PrimarySmtpAddress string }
		if json.Unmarshal(m, &o) == nil {
			out = append(out, firstOf(o.PrimarySmtpAddress, o.Name))
		}
	}
	return out
}

// oneMailbox runs Get-Mailbox for a resolved mailbox and decodes it.
func oneMailbox(env engine.Env, upn string, into any, sel ...string) error {
	items, err := exo(env, "Get-Mailbox", map[string]any{"Identity": upn}, sel...)
	if err != nil {
		return err
	}
	if len(items) == 0 {
		return fmt.Errorf("no mailbox found for %s", upn)
	}
	return json.Unmarshal(items[0], into)
}

// psAction wires a manifest to a PowerShell implementation.
func psAction(m engine.Manifest, impl engine.Impl) engine.Action {
	return engine.Action{Manifest: m, Impls: []engine.Impl{impl}}
}

func exchangePSActions() []engine.Action {
	mailbox := engine.Field{Name: "mailbox", Kind: engine.FieldUser, Required: true}
	user := engine.Field{Name: "user", Kind: engine.FieldUser, Required: true}
	addRemove := engine.Field{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"}
	perms := []string{"Exchange.ManageAsApp + Exchange Recipient Administrator"}
	return []engine.Action{
		psAction(engine.Manifest{ID: "mailbox.fullAccess", Page: "mail", Danger: engine.Write,
			Fields: []engine.Field{mailbox, user, addRemove}, Permissions: perms}, psFullAccess{}),
		psAction(engine.Manifest{ID: "mailbox.sendAs", Page: "mail", Danger: engine.Write,
			Fields: []engine.Field{mailbox, user, addRemove}, Permissions: perms}, psSendAs{}),
		psAction(engine.Manifest{ID: "mailbox.type", Page: "mail", Danger: engine.Write,
			Fields:      []engine.Field{mailbox, {Name: "type", Kind: engine.FieldChoice, Required: true, Options: []string{"shared", "regular"}, Default: "shared"}},
			Permissions: perms}, psMailboxType{}),
		psAction(engine.Manifest{ID: "mailbox.forwarding", Page: "mail", Danger: engine.Write,
			Fields: []engine.Field{mailbox,
				{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"set", "clear"}, Default: "set"},
				{Name: "forwardTo", Kind: engine.FieldUser},
				{Name: "keepCopy", Kind: engine.FieldChoice, Required: true, Options: []string{"yes", "no"}, Default: "yes"}},
			Permissions: perms}, psForwarding{}),
		psAction(engine.Manifest{ID: "mailbox.address", Page: "mail", Danger: engine.Write,
			Fields:      []engine.Field{mailbox, {Name: "address", Kind: engine.FieldText, Required: true}, addRemove},
			Permissions: perms}, psAddress{}),
		psAction(engine.Manifest{ID: "mailbox.calendarProcessing", Page: "mail", Danger: engine.Write,
			Fields:      []engine.Field{mailbox, {Name: "automate", Kind: engine.FieldChoice, Required: true, Options: []string{"AutoAccept", "AutoUpdate", "None"}, Default: "AutoAccept"}},
			Permissions: perms}, psCalendarProcessing{}),
		psAction(engine.Manifest{ID: "distributionList.membership", Page: "groups", Danger: engine.Write,
			Fields:      []engine.Field{user, {Name: "group", Kind: engine.FieldGroup, Required: true}, addRemove},
			Permissions: perms}, psDistributionList{}),
		psAction(engine.Manifest{ID: "transportRule.state", Page: "mail", Danger: engine.Write,
			Fields: []engine.Field{{Name: "rule", Kind: engine.FieldText, Required: true},
				{Name: "state", Kind: engine.FieldChoice, Required: true, Options: []string{"enabled", "disabled"}, Default: "disabled"}},
			Permissions: []string{"Exchange.ManageAsApp + Transport rules / Organization Management"}}, psTransportRule{}),
	}
}

// presence plans an add/remove of one grant from whether it exists now.
func presence(target, field, op, value string, has bool, ref map[string]string) engine.Change {
	ch := engine.Change{Target: target, Field: field, Op: op, Ref: ref}
	switch {
	case op == "add" && has:
		ch.Op, ch.Before, ch.After = "none", value, value
	case op == "add":
		ch.After = value
	case has:
		ch.Before = value
	default:
		ch.Op = "none"
	}
	return ch
}

type psBase struct{}

func (psBase) Backend() engine.Backend { return pwsh.BackendExchangePS }

// --- mailbox.fullAccess --------------------------------------------------------

type psFullAccess struct{ psBase }

func (psFullAccess) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	grantee, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	rows, err := exo(env, "Get-MailboxPermission", map[string]any{"Identity": owner.UPN, "User": grantee.UPN},
		"User", "AccessRights", "Deny", "IsInherited")
	if err != nil {
		return nil, err
	}
	has := false
	for _, raw := range rows {
		var r struct {
			AccessRights json.RawMessage
			Deny         bool
			IsInherited  bool
		}
		if json.Unmarshal(raw, &r) == nil && !r.Deny && !r.IsInherited && containsFold(strs(r.AccessRights), "FullAccess") {
			has = true
		}
	}
	return []engine.Change{presence(owner.UPN, "fullAccess", in["op"], grantee.UPN, has,
		map[string]string{"mailbox": owner.UPN, "user": grantee.UPN})}, nil
}

func (psFullAccess) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	params := map[string]any{"Identity": ch.Ref["mailbox"], "User": ch.Ref["user"], "AccessRights": "FullAccess", "InheritanceType": "All"}
	cmdlet := "Add-MailboxPermission"
	if ch.Op == "remove" {
		cmdlet, params["Confirm"] = "Remove-MailboxPermission", false
	} else {
		params["AutoMapping"] = true
	}
	_, err := exo(env, cmdlet, params)
	return err
}

// --- mailbox.sendAs ------------------------------------------------------------

type psSendAs struct{ psBase }

func (psSendAs) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	grantee, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	rows, err := exo(env, "Get-RecipientPermission", map[string]any{"Identity": owner.UPN, "Trustee": grantee.UPN, "AccessRights": "SendAs"},
		"Trustee", "AccessRights", "AccessControlType")
	if err != nil {
		return nil, err
	}
	has := false
	for _, raw := range rows {
		var r struct{ AccessControlType string }
		if json.Unmarshal(raw, &r) == nil && !strings.EqualFold(r.AccessControlType, "Deny") {
			has = true
		}
	}
	return []engine.Change{presence(owner.UPN, "sendAs", in["op"], grantee.UPN, has,
		map[string]string{"mailbox": owner.UPN, "user": grantee.UPN})}, nil
}

func (psSendAs) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	cmdlet := "Add-RecipientPermission"
	if ch.Op == "remove" {
		cmdlet = "Remove-RecipientPermission"
	}
	_, err := exo(env, cmdlet, map[string]any{"Identity": ch.Ref["mailbox"], "Trustee": ch.Ref["user"], "AccessRights": "SendAs", "Confirm": false})
	return err
}

// --- mailbox.type --------------------------------------------------------------

type psMailboxType struct{ psBase }

func (psMailboxType) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	var mb struct{ RecipientTypeDetails string }
	if err := oneMailbox(env, owner.UPN, &mb, "RecipientTypeDetails"); err != nil {
		return nil, err
	}
	before := map[string]string{"UserMailbox": "regular", "SharedMailbox": "shared"}[mb.RecipientTypeDetails]
	if before == "" {
		return nil, fmt.Errorf("%s is a %s; only user and shared mailboxes can be converted", owner.UPN, mb.RecipientTypeDetails)
	}
	ch := engine.Change{Target: owner.UPN, Field: "mailboxType", Op: "set", Before: before, After: in["type"],
		Ref: map[string]string{"mailbox": owner.UPN}}
	if before == in["type"] {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (psMailboxType) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	t := "Shared"
	if ch.After == "regular" {
		t = "Regular"
	}
	_, err := exo(env, "Set-Mailbox", map[string]any{"Identity": ch.Ref["mailbox"], "Type": t})
	return err
}

// --- mailbox.forwarding --------------------------------------------------------

type psForwarding struct{ psBase }

func (psForwarding) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	var mb struct {
		ForwardingSmtpAddress      string
		ForwardingAddress          string
		DeliverToMailboxAndForward bool
	}
	if err := oneMailbox(env, owner.UPN, &mb, "ForwardingSmtpAddress", "ForwardingAddress", "DeliverToMailboxAndForward"); err != nil {
		return nil, err
	}
	// ForwardingAddress (a recipient) wins over ForwardingSmtpAddress, so a
	// mailbox forwarding through it is "set" to anything new.
	smtp := strings.TrimPrefix(strings.TrimPrefix(mb.ForwardingSmtpAddress, "smtp:"), "SMTP:")
	current := firstOf(mb.ForwardingAddress, smtp)
	ref := map[string]string{"mailbox": owner.UPN}
	if in["op"] == "clear" {
		ch := engine.Change{Target: owner.UPN, Field: "forwarding", Op: "remove", Before: current, Ref: ref}
		if current == "" {
			ch.Op = "none"
		}
		return []engine.Change{ch}, nil
	}
	if in["forwardTo"] == "" {
		return nil, errors.New("choose who the mail is forwarded to")
	}
	to, err := lookupRecipient(env, in["forwardTo"])
	if err != nil {
		return nil, err
	}
	addr := firstOf(to.Mail, to.UPN)
	ref["to"] = addr
	keep := in["keepCopy"] == "yes"
	ch := engine.Change{Target: owner.UPN, Field: "forwarding", Op: "set", Before: current, After: addr, Ref: ref}
	if keep {
		ch.Note = "forwardKeepCopy"
	} else {
		ch.Note = "forwardNoCopy"
	}
	if mb.ForwardingAddress == "" && strings.EqualFold(smtp, addr) && mb.DeliverToMailboxAndForward == keep {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (psForwarding) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	params := map[string]any{"Identity": ch.Ref["mailbox"]}
	if ch.Op == "remove" {
		params["ForwardingSmtpAddress"] = nil
		params["ForwardingAddress"] = nil
		params["DeliverToMailboxAndForward"] = false
	} else {
		params["ForwardingSmtpAddress"] = "smtp:" + ch.Ref["to"]
		params["ForwardingAddress"] = nil // it would override the SMTP target
		params["DeliverToMailboxAndForward"] = in["keepCopy"] == "yes"
	}
	_, err := exo(env, "Set-Mailbox", params)
	return err
}

// --- mailbox.address ----------------------------------------------------------

type psAddress struct{ psBase }

var emailRe = regexp.MustCompile(`^[^@\s]+@[^@\s]+\.[^@\s]+$`)

func (psAddress) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	addr := strings.TrimSpace(strings.TrimPrefix(strings.TrimPrefix(in["address"], "smtp:"), "SMTP:"))
	if !emailRe.MatchString(addr) {
		return nil, errors.New("enter an email address, e.g. sales@contoso.com")
	}
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	var mb struct{ EmailAddresses json.RawMessage }
	if err := oneMailbox(env, owner.UPN, &mb, "EmailAddresses"); err != nil {
		return nil, err
	}
	has, primary := false, false
	for _, a := range strs(mb.EmailAddresses) {
		if strings.EqualFold(strings.TrimPrefix(strings.TrimPrefix(a, "smtp:"), "SMTP:"), addr) {
			has, primary = true, strings.HasPrefix(a, "SMTP:")
		}
	}
	if in["op"] == "remove" && primary {
		return nil, errors.New(addr + " is the primary address; make another address primary first")
	}
	return []engine.Change{presence(owner.UPN, "address", in["op"], addr, has,
		map[string]string{"mailbox": owner.UPN, "address": addr})}, nil
}

func (psAddress) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	verb := "Add"
	if ch.Op == "remove" {
		verb = "Remove"
	}
	_, err := exo(env, "Set-Mailbox", map[string]any{"Identity": ch.Ref["mailbox"],
		"EmailAddresses": map[string]any{verb: "smtp:" + ch.Ref["address"]}})
	return err
}

// --- mailbox.calendarProcessing -------------------------------------------------

type psCalendarProcessing struct{ psBase }

func (psCalendarProcessing) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	rows, err := exo(env, "Get-CalendarProcessing", map[string]any{"Identity": owner.UPN}, "AutomateProcessing")
	if err != nil {
		return nil, err
	}
	var cp struct{ AutomateProcessing string }
	if len(rows) > 0 {
		_ = json.Unmarshal(rows[0], &cp)
	}
	ch := engine.Change{Target: owner.UPN, Field: "calendarProcessing", Op: "set", Before: cp.AutomateProcessing, After: in["automate"],
		Ref: map[string]string{"mailbox": owner.UPN}}
	if strings.EqualFold(cp.AutomateProcessing, in["automate"]) {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (psCalendarProcessing) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	_, err := exo(env, "Set-CalendarProcessing", map[string]any{"Identity": ch.Ref["mailbox"], "AutomateProcessing": ch.After})
	return err
}

// --- distributionList.membership -------------------------------------------------

type psDistributionList struct{ psBase }

func (psDistributionList) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	member, err := lookupRecipient(env, in["user"])
	if err != nil {
		return nil, err
	}
	var g struct {
		Mail        string   `json:"mail"`
		DisplayName string   `json:"displayName"`
		MailEnabled bool     `json:"mailEnabled"`
		Unified     []string `json:"groupTypes"`
	}
	if err := env.Graph.Get(env.Ctx, "/groups/"+url.PathEscape(in["group"]), url.Values{"$select": {"mail,displayName,mailEnabled,groupTypes"}}, &g); err != nil {
		return nil, err
	}
	if !g.MailEnabled || g.Mail == "" || containsFold(g.Unified, "Unified") {
		return nil, fmt.Errorf("%s is not a distribution list or mail-enabled security group", firstOf(g.DisplayName, in["group"]))
	}
	rows, err := exo(env, "Get-DistributionGroupMember", map[string]any{"Identity": g.Mail, "ResultSize": "Unlimited"},
		"PrimarySmtpAddress", "ExternalDirectoryObjectId")
	if err != nil {
		return nil, err
	}
	has := false
	for _, raw := range rows {
		var m struct{ PrimarySmtpAddress, ExternalDirectoryObjectId string }
		if json.Unmarshal(raw, &m) == nil && (member.isAddress(m.PrimarySmtpAddress) || member.isAddress(m.ExternalDirectoryObjectId)) {
			has = true
		}
	}
	return []engine.Change{presence(member.UPN, "distributionList", in["op"], firstOf(g.DisplayName, g.Mail), has,
		map[string]string{"group": g.Mail, "user": firstOf(member.Mail, member.UPN)})}, nil
}

func (psDistributionList) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	cmdlet := "Add-DistributionGroupMember"
	params := map[string]any{"Identity": ch.Ref["group"], "Member": ch.Ref["user"], "BypassSecurityGroupManagerCheck": true}
	if ch.Op == "remove" {
		cmdlet, params["Confirm"] = "Remove-DistributionGroupMember", false
	}
	_, err := exo(env, cmdlet, params)
	return err
}

// --- transportRule.state -----------------------------------------------------------

type psTransportRule struct{ psBase }

func (psTransportRule) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	rows, err := exo(env, "Get-TransportRule", map[string]any{"Identity": in["rule"]}, "Name", "State")
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("no transport rule named %q", in["rule"])
	}
	var r struct{ Name, State string }
	if err := json.Unmarshal(rows[0], &r); err != nil || r.State == "" {
		return nil, fmt.Errorf("could not read the state of transport rule %q", in["rule"])
	}
	before := strings.ToLower(r.State)
	ch := engine.Change{Target: firstOf(r.Name, in["rule"]), Field: "ruleState", Op: "set", Before: before, After: in["state"],
		Ref: map[string]string{"rule": firstOf(r.Name, in["rule"])}}
	if before == in["state"] {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (psTransportRule) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	cmdlet := "Disable-TransportRule"
	if ch.After == "enabled" {
		cmdlet = "Enable-TransportRule"
	}
	_, err := exo(env, cmdlet, map[string]any{"Identity": ch.Ref["rule"], "Confirm": false})
	return err
}

// --- PowerShell stand-ins for the Admin API actions ---------------------------------

type psSendOnBehalf struct{ psBase }

func (psSendOnBehalf) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	owner, err := lookupRecipient(env, in["mailbox"])
	if err != nil {
		return nil, err
	}
	delegate, err := lookupRecipient(env, in["delegate"])
	if err != nil {
		return nil, err
	}
	var mb struct{ GrantSendOnBehalfTo json.RawMessage }
	if err := oneMailbox(env, owner.UPN, &mb, "GrantSendOnBehalfTo"); err != nil {
		return nil, err
	}
	has := false
	for _, d := range strs(mb.GrantSendOnBehalfTo) {
		if delegate.is(d) {
			has = true
		}
	}
	return []engine.Change{presence(owner.UPN, "sendOnBehalf", in["op"], delegate.UPN, has,
		map[string]string{"mailbox": owner.UPN, "delegate": firstOf(delegate.Mail, delegate.UPN)})}, nil
}

func (psSendOnBehalf) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	verb := "Add"
	if ch.Op == "remove" {
		verb = "Remove"
	}
	_, err := exo(env, "Set-Mailbox", map[string]any{"Identity": ch.Ref["mailbox"],
		"GrantSendOnBehalfTo": map[string]any{verb: ch.Ref["delegate"]}})
	return err
}

type psFolderPermission struct{ psBase }

func (psFolderPermission) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	if in["folder"] != "calendar" && calendarOnly[in["access"]] {
		return nil, errors.New(in["access"] + " exists only on calendar folders")
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
	items, err := exo(env, "Get-MailboxFolderPermission", map[string]any{"Identity": identity}, "User", "AccessRights")
	if err != nil {
		return nil, err
	}
	return planFolderChange(items, grantee, identity, owner.UPN, in["access"]), nil
}

func (psFolderPermission) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	params := map[string]any{"Identity": ch.Ref["identity"], "User": ch.Ref["user"]}
	cmdlet := "Remove-MailboxFolderPermission"
	switch ch.Op {
	case "add":
		cmdlet, params["AccessRights"] = "Add-MailboxFolderPermission", ch.After
	case "set":
		cmdlet, params["AccessRights"] = "Set-MailboxFolderPermission", ch.After
	default:
		params["Confirm"] = false
	}
	_, err := exo(env, cmdlet, params)
	return err
}

func containsFold(list []string, v string) bool {
	for _, x := range list {
		if strings.EqualFold(x, v) {
			return true
		}
	}
	return false
}
