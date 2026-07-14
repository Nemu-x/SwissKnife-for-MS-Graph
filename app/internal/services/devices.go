package services

import (
	"encoding/json"
	"net/url"

	"swissknife-app/internal/session"
)

// DevicesService — Entra ID (directory) devices, distinct from Intune managed
// devices. Also exposes BitLocker recovery keys.
type DevicesService struct {
	s *session.Session
}

func NewDevicesService(s *session.Session) *DevicesService { return &DevicesService{s: s} }

func (d *DevicesService) List(search string, maxItems int) ([]json.RawMessage, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	params := url.Values{
		"$top":    {"100"},
		"$select": {"id,deviceId,displayName,operatingSystem,operatingSystemVersion,accountEnabled,isCompliant,isManaged,trustType,approximateLastSignInDateTime"},
	}
	if search != "" {
		params.Set("$filter", "startswith(displayName,'"+escapeODataLiteral(search)+"')")
	}
	return c.ListAll(d.s.Ctx(), "/devices", params, maxItems)
}

func (d *DevicesService) Get(deviceID string) (json.RawMessage, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(d.s.Ctx(), "/devices/"+url.PathEscape(deviceID), nil, &out)
	return out, err
}

func (d *DevicesService) setEnabled(deviceID string, enabled bool) error {
	if err := d.s.GuardWrite(); err != nil {
		return err
	}
	c, err := d.s.Client()
	if err != nil {
		return err
	}
	action := "devices.disable"
	if enabled {
		action = "devices.enable"
	}
	err = c.Patch(d.s.Ctx(), "/devices/"+url.PathEscape(deviceID), map[string]any{"accountEnabled": enabled}, nil)
	d.s.Record(action, deviceID, "", err)
	return err
}

func (d *DevicesService) Enable(deviceID string) error  { return d.setEnabled(deviceID, true) }
func (d *DevicesService) Disable(deviceID string) error { return d.setEnabled(deviceID, false) }

// Delete removes a device from the directory. Destructive: typed confirm.
func (d *DevicesService) Delete(deviceID, confirm string) error {
	if err := d.s.GuardDestructive(deviceID, confirm); err != nil {
		return err
	}
	c, err := d.s.Client()
	if err != nil {
		return err
	}
	err = c.Delete(d.s.Ctx(), "/devices/"+url.PathEscape(deviceID))
	d.s.Record("devices.delete", deviceID, "", err)
	return err
}

// BitLockerKeys lists BitLocker recovery key metadata (id, deviceId, volumeType).
// The key value itself is fetched separately via BitLockerKey.
func (d *DevicesService) BitLockerKeys(maxItems int) ([]json.RawMessage, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(d.s.Ctx(), "/informationProtection/bitlocker/recoveryKeys", nil, maxItems)
}

// BitLockerKey returns a single recovery key including its secret value.
// Requires BitlockerKey.Read.All; reading the value is audited.
func (d *DevicesService) BitLockerKey(keyID string) (json.RawMessage, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(d.s.Ctx(), "/informationProtection/bitlocker/recoveryKeys/"+url.PathEscape(keyID),
		url.Values{"$select": {"key,deviceId,volumeType,createdDateTime"}}, &out)
	d.s.Record("devices.bitlockerKey", keyID, "", err)
	return out, err
}
