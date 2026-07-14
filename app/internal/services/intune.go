package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

type IntuneService struct {
	s *session.Session
}

func NewIntuneService(s *session.Session) *IntuneService { return &IntuneService{s: s} }

func (i *IntuneService) Devices(maxItems int) ([]json.RawMessage, error) {
	c, err := i.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(i.s.Ctx(), "/deviceManagement/managedDevices", topParams(100), maxItems)
}

func (i *IntuneService) Device(deviceID string) (json.RawMessage, error) {
	c, err := i.s.Client()
	if err != nil {
		return nil, err
	}
	var out json.RawMessage
	err = c.Get(i.s.Ctx(), "/deviceManagement/managedDevices/"+url.PathEscape(deviceID), nil, &out)
	return out, err
}

// Wipe is destructive: typed confirm on the device id.
func (i *IntuneService) Wipe(deviceID string, keepEnrollmentData, keepUserData bool, confirm string) error {
	if err := i.s.GuardDestructive(deviceID, confirm); err != nil {
		return err
	}
	c, err := i.s.Client()
	if err != nil {
		return err
	}
	body := map[string]any{
		"keepEnrollmentData": keepEnrollmentData,
		"keepUserData":       keepUserData,
		"macOsUnlockCode":    nil,
	}
	err = c.Post(i.s.Ctx(), "/deviceManagement/managedDevices/"+url.PathEscape(deviceID)+"/wipe", body, nil)
	i.s.Record("intune.wipe", deviceID,
		"keepEnrollment="+strconv.FormatBool(keepEnrollmentData)+" keepUser="+strconv.FormatBool(keepUserData), err)
	return err
}

func (i *IntuneService) Retire(deviceID, confirm string) error {
	if err := i.s.GuardDestructive(deviceID, confirm); err != nil {
		return err
	}
	c, err := i.s.Client()
	if err != nil {
		return err
	}
	err = c.Post(i.s.Ctx(), "/deviceManagement/managedDevices/"+url.PathEscape(deviceID)+"/retire", map[string]any{}, nil)
	i.s.Record("intune.retire", deviceID, "", err)
	return err
}

func (i *IntuneService) RemoteLock(deviceID, confirm string) error {
	if err := i.s.GuardDestructive(deviceID, confirm); err != nil {
		return err
	}
	c, err := i.s.Client()
	if err != nil {
		return err
	}
	err = c.Post(i.s.Ctx(), "/deviceManagement/managedDevices/"+url.PathEscape(deviceID)+"/remoteLock", map[string]any{}, nil)
	i.s.Record("intune.remoteLock", deviceID, "", err)
	return err
}
