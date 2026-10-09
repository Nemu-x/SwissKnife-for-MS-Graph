package services

import (
	"errors"
	"fmt"
	"net/url"
	"strings"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// Restoring configuration from a snapshot: the plan is the difference
// between the snapshot and the tenant now, per object, and applying it puts
// the snapshot's values back (or recreates what was deleted).

// restorable describes a Graph snapshot section that can be written back.
type restorable struct {
	path string // collection path
	// writable lists the properties sent on PATCH/POST; nil sends every
	// property except restoreSkip (Intune's many policy types).
	writable []string
	typed    bool // @odata.type must accompany writes
	// assign: assignments are restored too, through {path}/{id}/assign.
	assign bool
	// create holds properties a POST needs that the snapshot does not keep.
	create map[string]any
}

// restoreSkip are read-only or separately restored properties of objects
// restored with every property.
var restoreSkip = map[string]bool{"id": true, "createdDateTime": true, "lastModifiedDateTime": true, "version": true,
	"assignments": true, "supportsScopeTags": true, "scheduledActionsForRule": true, "@odata.type": true,
	// Windows Update rings: set by Intune as pauses expire and rollbacks run.
	"qualityUpdatesPauseExpiryDateTime": true, "featureUpdatesPauseExpiryDateTime": true,
	"qualityUpdatesWillBeRolledBack": true, "featureUpdatesWillBeRolledBack": true,
	"qualityUpdatesRollbackStartDateTime": true, "featureUpdatesRollbackStartDateTime": true,
	// Read back masked: writing it would replace the real key with the mask.
	"productKey": true}

var restorableSections = map[string]restorable{
	"conditionalAccessPolicies": {path: "/identity/conditionalAccess/policies",
		writable: []string{"displayName", "state", "conditions", "grantControls", "sessionControls"}},
	"namedLocations": {path: "/identity/conditionalAccess/namedLocations",
		writable: []string{"displayName", "ipRanges", "isTrusted", "countriesAndRegions", "includeUnknownCountriesAndRegions", "countryLookupMethod"}, typed: true},
	"intuneConfigurations":     {path: "/deviceManagement/deviceConfigurations", typed: true, assign: true},
	"intuneCompliancePolicies": {path: "/deviceManagement/deviceCompliancePolicies", typed: true, assign: true,
		// A compliance policy cannot be created without the action taken on
		// non-compliance: the snapshot's (taken since this version), else
		// Intune's default — non-compliant at once.
		create: map[string]any{"scheduledActionsForRule": []any{map[string]any{"ruleName": "PasswordRequired",
			"scheduledActionConfigurations": []any{map[string]any{"actionType": "block", "gracePeriodHours": 0}}}}}},
}

// psRestore describes a PowerShell snapshot section that can be written back
// with its Set- cmdlet (and New- for a deleted object, when it can be made
// from what the snapshot keeps).
type psRestore struct {
	set    string
	create string   // "" = a deleted object cannot be recreated
	props  []string // properties passed back to set/create
	global bool     // one object, no -Identity (organization config)
	state  bool     // State goes through Enable-/Disable-TransportRule
}

var psRestorable = map[string]psRestore{
	"exchangeOrganizationConfig": {set: "Set-OrganizationConfig", global: true, props: []string{"AuditDisabled", "OAuth2ClientProfileEnabled",
		"CustomerLockBoxEnabled", "MailTipsExternalRecipientsTipsEnabled", "DefaultAuthenticationPolicy", "FocusedInboxOn", "PublicFoldersEnabled"}},
	// Priority is not written back: each change shifts the other rules, and
	// rules deleted since make old priorities invalid.
	"transportRules": {set: "Set-TransportRule", props: []string{"Name", "Mode"}, state: true},
	"teamsMeetingPolicies": {set: "Set-CsTeamsMeetingPolicy", create: "New-CsTeamsMeetingPolicy", props: []string{"AllowCloudRecording",
		"AllowTranscription", "AllowAnonymousUsersToJoinMeeting", "AutoAdmittedUsers", "AllowExternalParticipantGiveRequestControl", "AllowMeetNow"}},
}

// restoreOrder is the order "all" restores in: named locations before the
// policies that point at them.
var restoreOrder = []string{"namedLocations", "conditionalAccessPolicies", "intuneConfigurations", "intuneCompliancePolicies",
	"exchangeOrganizationConfig", "transportRules", "teamsMeetingPolicies"}

func snapshotRestoreAction(s *session.Session) engine.Action {
	return engine.Action{
		Manifest: engine.Manifest{
			ID: "config.restore", Capability: "config.snapshot.restore", Page: "security", Danger: engine.Destructive,
			Fields: []engine.Field{
				{Name: "snapshot", Kind: engine.FieldSnapshot, Required: true},
				{Name: "section", Kind: engine.FieldChoice, Required: true, Options: append(append([]string{}, restoreOrder...), "all"), Default: "conditionalAccessPolicies"},
				{Name: "object", Kind: engine.FieldText},
			},
			ConfirmField: "snapshot",
			Permissions: []string{"Policy.ReadWrite.ConditionalAccess", "Policy.Read.All", "DeviceManagementConfiguration.ReadWrite.All",
				"Exchange.ManageAsApp + Organization Management", "Teams Administrator"},
		},
		Impls: []engine.Impl{snapshotRestore{svc: NewSnapshotService(s)}},
	}
}

type snapshotRestore struct{ svc *SnapshotService }

func (snapshotRestore) Backend() engine.Backend { return engine.BackendGraph }

// subset keeps the writable properties of an object.
func (r restorable) subset(o map[string]any) map[string]any {
	out := map[string]any{}
	if r.writable == nil {
		for k, v := range o {
			// Secrets read back as null: never write the null over them.
			if !restoreSkip[k] && v != nil {
				out[k] = v
			}
		}
	}
	for _, k := range r.writable {
		if v, ok := o[k]; ok {
			out[k] = v
		}
	}
	// An authentication strength is read back in full but written by id only.
	if g, ok := out["grantControls"].(map[string]any); ok {
		if as, ok := g["authenticationStrength"].(map[string]any); ok {
			g2 := make(map[string]any, len(g))
			for k, v := range g {
				g2[k] = v
			}
			g2["authenticationStrength"] = map[string]any{"id": as["id"]}
			out["grantControls"] = g2
		}
	}
	if r.typed {
		if t, ok := o["@odata.type"]; ok {
			out["@odata.type"] = t
		}
	}
	return out
}

// sameTenant reports whether a snapshot was taken in the connected tenant.
// Snapshots from before tenant ids were recorded never match: profile names
// are local labels and prove nothing about the directory.
func sameTenant(m SnapshotMeta, tenantID string) error {
	if m.TenantID == "" {
		return errors.New("this snapshot does not record its tenant (taken by an older version); take a new snapshot")
	}
	if tenantID == "" || !strings.EqualFold(m.TenantID, tenantID) {
		return fmt.Errorf("the snapshot was taken in another tenant (%s)", m.Tenant)
	}
	return nil
}

// liveByName indexes a collection's current objects by lower-cased display
// name; a name can belong to several objects.
func liveByName(env engine.Env, path string) (map[string][]string, error) {
	objs, err := graphapi.ListAllInto[map[string]any](env.Ctx, env.Graph, path, url.Values{"$select": {"id,displayName"}}, 0)
	if err != nil {
		return nil, err
	}
	out := map[string][]string{}
	for _, o := range objs {
		name, _ := o["displayName"].(string)
		id, _ := o["id"].(string)
		if name != "" && id != "" {
			out[strings.ToLower(name)] = append(out[strings.ToLower(name)], id)
		}
	}
	return out, nil
}

// remapLocations points a policy's location conditions at today's named
// locations: a location recreated after deletion has a new id, so the
// snapshot's id is matched through its name (only when exactly one location
// has that name). Unresolvable ids are returned.
func remapLocations(policy map[string]any, snapNames map[string]string, live map[string][]string, liveIDs map[string]bool) []string {
	cond, _ := policy["conditions"].(map[string]any)
	locs, _ := cond["locations"].(map[string]any)
	if locs == nil {
		return nil
	}
	var missing []string
	for _, k := range []string{"includeLocations", "excludeLocations"} {
		list, _ := locs[k].([]any)
		for i, v := range list {
			id, _ := v.(string)
			if id == "" || id == "All" || id == "AllTrusted" || liveIDs[id] {
				continue
			}
			if ids := live[strings.ToLower(snapNames[id])]; len(ids) == 1 && snapNames[id] != "" {
				list[i] = ids[0]
				continue
			}
			missing = append(missing, id)
		}
	}
	return missing
}

