package actions

import (
	"encoding/json"
	"errors"
	"fmt"
	"strings"

	"swissknife-app/internal/engine"
)

// Exchange Online Protection through PowerShell: block a sender in the
// Tenant Allow/Block List, and look into / release from quarantine.

var securityCmdlets = []string{
	"Get-TenantAllowBlockListItems", "New-TenantAllowBlockListItems", "Remove-TenantAllowBlockListItems",
	"Get-QuarantineMessage", "Release-QuarantineMessage",
}

func securityPSActions() []engine.Action {
	perms := []string{"Exchange.ManageAsApp + Security Administrator (or Exchange Administrator)"}
	return []engine.Action{
		psAction(engine.Manifest{ID: "mail.blockSender", Page: "security", Danger: engine.Write,
			Fields: []engine.Field{
				{Name: "sender", Kind: engine.FieldText, Required: true},
				{Name: "op", Kind: engine.FieldChoice, Required: true, Options: []string{"add", "remove"}, Default: "add"},
			}, Permissions: perms}, psBlockSender{}),
		{Manifest: engine.Manifest{ID: "mail.quarantine", Page: "security", Danger: engine.Read,
			Fields: []engine.Field{
				{Name: "recipient", Kind: engine.FieldUser},
				{Name: "sender", Kind: engine.FieldText},
			}, Permissions: perms},
			Impls: []engine.Impl{engine.ReadImpl(psQuarantineList{})}},
		psAction(engine.Manifest{ID: "mail.releaseQuarantine", Page: "security", Danger: engine.Write,
			Fields:      []engine.Field{{Name: "identity", Kind: engine.FieldText, Required: true}},
			Permissions: perms}, psReleaseQuarantine{}),
	}
}

// --- mail.blockSender -------------------------------------------------------------

type psBlockSender struct{ psBase }

func blockEntry(in string) (string, error) {
	e := strings.ToLower(strings.TrimSpace(in))
	if !senderAddrRe.MatchString(e) && !senderDomRe.MatchString(e) {
		return "", errors.New("enter a sender address (bad@example.com) or a domain (example.com)")
	}
	return e, nil
}

func (psBlockSender) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	entry, err := blockEntry(in["sender"])
	if err != nil {
		return nil, err
	}
	// Blocking a whole domain of the tenant, or a free-mail provider, would
	// cut off far more than one attacker.
	if !strings.Contains(entry, "@") && in["op"] == "add" {
		internal, err := tenantDomains(env)
		if err != nil {
			return nil, err
		}
		if internal[entry] || freeMail[entry] {
			return nil, fmt.Errorf("refusing to block the whole domain %s — block the sender's address instead", entry)
		}
	}
	rows, err := exo(env, "Get-TenantAllowBlockListItems", map[string]any{"ListType": "Sender", "Entry": entry}, "Value", "Action")
	if err != nil && !strings.Contains(strings.ToLower(err.Error()), "not found") {
		return nil, err
	}
	blocked := false
	for _, raw := range rows {
		var r struct{ Value, Action string }
		if json.Unmarshal(raw, &r) != nil || !strings.EqualFold(r.Value, entry) {
			continue
		}
		if strings.EqualFold(r.Action, "Allow") && in["op"] == "add" {
			return nil, fmt.Errorf("%s is on the allow list — remove that entry first", entry)
		}
		if strings.EqualFold(r.Action, "Block") {
			blocked = true
		}
	}
	return []engine.Change{presence("Tenant Allow/Block List", "blockedSender", in["op"], entry, blocked,
		map[string]string{"entry": entry})}, nil
}

func (psBlockSender) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	if ch.Op == "remove" {
		_, err := exo(env, "Remove-TenantAllowBlockListItems", map[string]any{"ListType": "Sender", "Entries": []string{ch.Ref["entry"]}})
		return err
	}
	_, err := exo(env, "New-TenantAllowBlockListItems", map[string]any{
		"ListType": "Sender", "Block": true, "Entries": []string{ch.Ref["entry"]}, "NoExpiration": true,
		"Notes": "Blocked by SwissKnife",
	})
	return err
}

// --- mail.quarantine ---------------------------------------------------------------

type psQuarantineList struct{ psBase }

func (psQuarantineList) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	params := map[string]any{"PageSize": 100}
	if in["recipient"] != "" {
		r, err := lookupRecipient(env, in["recipient"])
		if err != nil {
			return nil, err
		}
		params["RecipientAddress"] = []string{firstOf(r.Mail, r.UPN)}
	}
	if s := strings.TrimSpace(in["sender"]); s != "" {
		params["SenderAddress"] = []string{s}
	}
	rows, err := exo(env, "Get-QuarantineMessage", params,
		"Identity", "ReceivedTime", "SenderAddress", "RecipientAddress", "Subject", "QuarantineTypes", "ReleaseStatus")
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: []string{"received", "sender", "recipient", "subject", "type", "status", "identity"}}
	for _, raw := range rows {
		var m struct {
			Identity, ReceivedTime, SenderAddress, Subject, QuarantineTypes, ReleaseStatus string
			RecipientAddress                                                               json.RawMessage
		}
		if json.Unmarshal(raw, &m) != nil {
			continue
		}
		res.Rows = append(res.Rows, engine.Row{
			"received": m.ReceivedTime, "sender": m.SenderAddress, "recipient": strings.Join(strs(m.RecipientAddress), ", "),
			"subject": m.Subject, "type": m.QuarantineTypes, "status": m.ReleaseStatus, "identity": m.Identity,
		})
	}
	return res, nil
}

// --- mail.releaseQuarantine -----------------------------------------------------------

type psReleaseQuarantine struct{ psBase }

func (psReleaseQuarantine) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	id := strings.TrimSpace(in["identity"])
	rows, err := exo(env, "Get-QuarantineMessage", map[string]any{"Identity": id}, "Subject", "SenderAddress", "ReleaseStatus")
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("no quarantined message with identity %s", id)
	}
	var m struct{ Subject, SenderAddress, ReleaseStatus string }
	_ = json.Unmarshal(rows[0], &m)
	ch := engine.Change{Target: fmt.Sprintf("%q from %s", m.Subject, m.SenderAddress), Field: "quarantine", Op: "set",
		Before: m.ReleaseStatus, After: "released", Ref: map[string]string{"identity": id}}
	if strings.EqualFold(m.ReleaseStatus, "Released") {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

func (psReleaseQuarantine) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	// The cmdlet needs -ReleaseToAll or -User; ReleaseToAll means the
	// message's original recipients.
	_, err := exo(env, "Release-QuarantineMessage", map[string]any{"Identity": ch.Ref["identity"], "ReleaseToAll": true})
	return err
}
