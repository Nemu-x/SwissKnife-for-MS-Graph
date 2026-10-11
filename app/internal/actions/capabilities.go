package actions

import "swissknife-app/internal/engine"

// capabilities are the stable public names of the built-in actions
// (<service>.<object>.<verb>). Workflows and hub packs may name steps by
// them, so a name, once published, is never changed or reused.
var capabilities = map[string]string{
	"user.signIn":                 "entra.user.signIn",
	"user.revokeSessions":         "entra.user.revokeSessions",
	"user.manager":                "entra.user.manager",
	"user.usageLocation":          "entra.user.usageLocation",
	"user.resetMfa":               "entra.user.resetMfa",
	"user.resetPassword":          "entra.user.resetPassword",
	"group.membership":            "entra.group.membership",
	"license.assign":              "entra.license.assign",
	"mailbox.sendOnBehalf":        "exchange.mailbox.sendOnBehalf",
	"mailbox.folderPermission":    "exchange.mailbox.folderPermission",
	"mailbox.fullAccess":          "exchange.mailbox.fullAccess",
	"mailbox.sendAs":              "exchange.mailbox.sendAs",
	"mailbox.type":                "exchange.mailbox.type",
	"mailbox.forwarding":          "exchange.mailbox.forwarding",
	"mailbox.address":             "exchange.mailbox.address",
	"mailbox.calendarProcessing":  "exchange.mailbox.calendarProcessing",
	"mailbox.statistics":          "exchange.mailbox.statistics",
	"distributionList.membership": "exchange.distributionList.membership",
	"transportRule.state":         "exchange.transportRule.state",
	"mail.ruleAudit":              "exchange.inboxRule.audit",
	"mail.disableRules":           "exchange.inboxRule.disable",
	"mail.blockSender":            "defender.sender.block",
	"mail.quarantine":             "defender.quarantine.list",
	"mail.releaseQuarantine":      "defender.quarantine.release",
	"mail.reportThreat":           "defender.mail.reportThreat",
	"mail.purge":                  "compliance.mail.purge",
	"teams.userPolicy":            "teams.policy.assignUser",
	"teams.groupPolicy":           "teams.policy.assignGroup",
	"teams.policies":              "teams.policy.list",
	"teams.effectivePolicies":     "teams.policy.effective",
	"report.inactiveUsers":        "reports.users.inactive",
	"report.licenseWaste":         "reports.licenses.waste",
	"report.guests":               "reports.users.guests",
	"report.mfaStatus":            "reports.users.mfaStatus",
	"report.privilegedRoles":      "reports.roles.privileged",
	"report.mailboxSizes":         "reports.mailbox.sizes",
	"report.mailboxForwarding":    "reports.mailbox.forwarding",
	"intune.assignedTo":           "intune.policy.assignedTo",
	"intune.unassigned":           "intune.policy.unassigned",
	"intune.conflicts":            "intune.policy.conflicts",
	"intune.assign":               "intune.policy.assign",
	"ad.findUser":                 "ad.user.find",
	"ad.userState":                "ad.user.state",
	"ad.unlock":                   "ad.user.unlock",
	"ad.resetPassword":            "ad.user.resetPassword",
	"ad.groupMembership":          "ad.group.membership",
}

func withCapabilities(list []engine.Action) []engine.Action {
	for i := range list {
		list[i].Capability = capabilities[list[i].ID]
	}
	return list
}
