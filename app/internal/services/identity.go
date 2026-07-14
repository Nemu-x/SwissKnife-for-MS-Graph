package services

import (
	"encoding/json"
	"errors"
	"net/url"
	"strings"
)

// Identity operations on UsersService: lifecycle (create/update/delete),
// manager, usage location, and deleted-user recovery.

// CreateUser creates a cloud user. usageLocation is optional but recommended
// (licensing needs it). Returns the created user object.
func (u *UsersService) CreateUser(displayName, upn, mailNickname, password string, forceChange bool, usageLocation string) (json.RawMessage, error) {
	if err := u.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	body := map[string]any{
		"accountEnabled":    true,
		"displayName":       displayName,
		"mailNickname":      mailNickname,
		"userPrincipalName": upn,
		"passwordProfile": map[string]any{
			"password":                      password,
			"forceChangePasswordNextSignIn": forceChange,
		},
	}
	if usageLocation != "" {
		body["usageLocation"] = usageLocation
	}
	var out json.RawMessage
	err = c.Post(u.s.Ctx(), "/users", body, &out)
	u.s.Record("users.create", upn, "", err)
	return out, err
}

// Update patches arbitrary user fields from a JSON object (e.g. jobTitle,
// department, officeLocation). Not destructive, but write-guarded.
func (u *UsersService) Update(user, patchJSON string) error {
	if err := u.s.GuardWrite(); err != nil {
		return err
	}
	var body map[string]any
	if err := json.Unmarshal([]byte(patchJSON), &body); err != nil {
		return errors.New("patch is not valid JSON: " + err.Error())
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	err = c.Patch(u.s.Ctx(), "/users/"+url.PathEscape(user), body, nil)
	u.s.Record("users.update", user, keysOf(body), err)
	return err
}

// SetUsageLocation sets the ISO 3166-1 alpha-2 country code (prerequisite for licensing).
func (u *UsersService) SetUsageLocation(user, location string) error {
	if err := u.s.GuardWrite(); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	err = c.Patch(u.s.Ctx(), "/users/"+url.PathEscape(user), map[string]any{"usageLocation": strings.ToUpper(location)}, nil)
	u.s.Record("users.setUsageLocation", user, "loc="+location, err)
	return err
}

// Delete removes a user (soft-delete — recoverable for 30 days). Destructive.
func (u *UsersService) Delete(user, confirm string) error {
	if err := u.s.GuardDestructive(user, confirm); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	err = c.Delete(u.s.Ctx(), "/users/"+url.PathEscape(user))
	u.s.Record("users.delete", user, "", err)
	return err
}

func (u *UsersService) GetManager(user string) (json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(u.s.Ctx(), "/users/"+url.PathEscape(user)+"/manager", nil, &out)
	return out, err
}

// SetManager assigns manager (by UPN/id) to user.
func (u *UsersService) SetManager(user, managerUpn string) error {
	if err := u.s.GuardWrite(); err != nil {
		return err
	}
	c, err := u.s.Client()
	if err != nil {
		return err
	}
	var mgr struct {
		ID string `json:"id"`
	}
	if err := c.Get(u.s.Ctx(), "/users/"+url.PathEscape(managerUpn), url.Values{"$select": {"id"}}, &mgr); err != nil {
		return err
	}
	body := map[string]any{"@odata.id": "https://graph.microsoft.com/v1.0/users/" + mgr.ID}
	err = c.Put(u.s.Ctx(), "/users/"+url.PathEscape(user)+"/manager/$ref", body, nil)
	u.s.Record("users.setManager", user, "manager="+managerUpn, err)
	return err
}

// ListDeleted lists soft-deleted users (recoverable for 30 days).
func (u *UsersService) ListDeleted(maxItems int) ([]json.RawMessage, error) {
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(u.s.Ctx(), "/directory/deletedItems/microsoft.graph.user", topParams(100), maxItems)
}

// RestoreDeleted restores a soft-deleted directory object (user/group) by id.
func (u *UsersService) RestoreDeleted(objectID string) (json.RawMessage, error) {
	if err := u.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := u.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Post(u.s.Ctx(), "/directory/deletedItems/"+url.PathEscape(objectID)+"/restore", map[string]any{}, &out)
	u.s.Record("users.restore", objectID, "", err)
	return out, err
}

func keysOf(m map[string]any) string {
	ks := make([]string, 0, len(m))
	for k := range m {
		ks = append(ks, k)
	}
	return strings.Join(ks, ",")
}
