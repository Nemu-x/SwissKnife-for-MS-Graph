package services

import (
	"encoding/json"
	"errors"
	"strings"

	"swissknife-app/internal/graphapi"
)

// OpError is the structured error envelope that crosses the Wails boundary.
// Wails v2 serializes Go errors as bare strings, so the envelope is JSON
// encoded into the string behind an "operr:" prefix; the frontend detects the
// prefix and parses the payload (lib/graphError.ts).
type OpError struct {
	Code      string `json:"code,omitempty"`
	Status    int    `json:"status,omitempty"`
	RequestID string `json:"requestId,omitempty"`
	Hint      string `json:"hint,omitempty"` // missing Graph application permission, when known
	Message   string `json:"message"`
}

func (e *OpError) Error() string {
	b, _ := json.Marshal(e)
	return "operr:" + string(b)
}

// wrapOpErr converts an error into the envelope, enriching Graph errors with
// code/status/requestId and — for 403s on known endpoint families — the
// missing-permission hint. Non-Graph errors pass through unchanged (their
// plain text is already the best representation).
func wrapOpErr(err error) error {
	if err == nil {
		return nil
	}
	var ge *graphapi.GraphError
	if !errors.As(err, &ge) {
		return err
	}
	oe := &OpError{Code: ge.Code, Status: ge.StatusCode, RequestID: ge.RequestID, Message: ge.Message}
	if ge.StatusCode == 403 {
		oe.Hint = permissionHint(ge.Path)
	}
	return oe
}

// permHints maps Graph endpoint families (path substrings) to the application
// permission a 403 most likely means is missing. First match wins; unknown
// endpoints get no hint (no wrong guesses).
var permHints = []struct{ needle, perm string }{
	{"/messageRules", "Mail.ReadWrite"},
	{"/mailFolders", "Mail.ReadWrite"},
	{"/mailboxSettings", "MailboxSettings.ReadWrite"},
	{"/calendarPermissions", "Calendars.ReadWrite"},
	{"/calendar", "Calendars.ReadWrite"},
	{"/assignLicense", "User.ReadWrite.All"},
	{"/licenseDetails", "User.Read.All"},
	{"/revokeSignInSessions", "User.ReadWrite.All"},
	{"/authentication/", "UserAuthenticationMethod.ReadWrite.All"},
	{"/managedDevices", "DeviceManagementManagedDevices.ReadWrite.All"},
	{"/drive", "Files.ReadWrite.All"},
	{"/invitations", "User.Invite.All"},
	{"/applications", "Application.ReadWrite.All"},
	{"/groups", "GroupMember.ReadWrite.All"},
	{"/teams", "TeamMember.ReadWrite.All"},
	{"/users", "User.ReadWrite.All"},
}

// permissionHint returns the likely missing application permission for a Graph
// endpoint path, or "" when unknown.
func permissionHint(path string) string {
	for _, h := range permHints {
		if strings.Contains(path, h.needle) {
			return h.perm
		}
	}
	return ""
}
