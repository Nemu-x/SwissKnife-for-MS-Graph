package services

import (
	"encoding/json"
	"errors"

	wrt "github.com/wailsapp/wails/v2/pkg/runtime"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/session"
)

// ConnectService manages connection profiles and the current session.
type ConnectService struct {
	s     *session.Session
	store *secrets.Store
}

func NewConnectService(s *session.Session, store *secrets.Store) *ConnectService {
	return &ConnectService{s: s, store: store}
}

func (c *ConnectService) Profiles() ([]secrets.Profile, error) {
	list, err := c.store.List()
	if list == nil {
		list = []secrets.Profile{}
	}
	return list, err
}

// SaveProfile: secret == "" keeps the stored secret unchanged.
func (c *ConnectService) SaveProfile(p secrets.Profile, secret string) (secrets.Profile, error) {
	if p.Name == "" || p.TenantID == "" || p.ClientID == "" {
		return secrets.Profile{}, errors.New("name, tenantId and clientId are required")
	}
	if p.AuthMode == "" {
		p.AuthMode = string(auth.ModeClientSecret)
	}
	return c.store.Save(p, secret)
}

func (c *ConnectService) DeleteProfile(profileID string) error {
	return c.store.Delete(profileID)
}

// ConnectRequest holds connection parameters: either ProfileID (secret from keychain),
// or ad-hoc credentials (Secret in memory; RememberAs != "" saves a profile).
type ConnectRequest struct {
	ProfileID  string `json:"profileId"`
	TenantID   string `json:"tenantId"`
	ClientID   string `json:"clientId"`
	Secret     string `json:"secret"`
	AuthMode   string `json:"authMode"` // client_secret | device_code
	RememberAs string `json:"rememberAs"`
}

// Status is the session state exposed to the frontend.
type Status struct {
	Connected   bool            `json:"connected"`
	ProfileName string          `json:"profileName"`
	ReadOnly    bool            `json:"readOnly"`
	Org         json.RawMessage `json:"org,omitempty"`
}

// Connect establishes the session and runs a self-test GET /organization.
func (c *ConnectService) Connect(req ConnectRequest) (*Status, error) {
	var (
		provider *auth.TokenProvider
		name     string
		err      error
	)

	mode := req.AuthMode
	tenant, client, secret := req.TenantID, req.ClientID, req.Secret

	if req.ProfileID != "" {
		profiles, perr := c.store.List()
		if perr != nil {
			return nil, perr
		}
		var found *secrets.Profile
		for i := range profiles {
			if profiles[i].ID == req.ProfileID {
				found = &profiles[i]
				break
			}
		}
		if found == nil {
			return nil, errors.New("profile not found")
		}
		tenant, client, mode, name = found.TenantID, found.ClientID, found.AuthMode, found.Name
		if mode == string(auth.ModeClientSecret) {
			secret, err = c.store.Secret(found.ID)
			if err != nil {
				return nil, err
			}
		}
	}

	switch mode {
	case string(auth.ModeDeviceCode):
		provider, err = auth.NewDeviceCode(tenant, client, func(url, code, msg string) {
			wrt.EventsEmit(c.s.Ctx(), "auth:deviceCode", map[string]string{
				"url": url, "code": code, "message": msg,
			})
		})
	default:
		if secret == "" {
			return nil, errors.New("client secret is required")
		}
		provider, err = auth.NewClientSecret(tenant, client, secret)
	}
	if err != nil {
		return nil, err
	}

	gc := graphapi.New(provider)

	// self-test plus token warm-up (device code: the user enters the code here)
	var org json.RawMessage
	if err := gc.Get(c.s.Ctx(), "/organization", nil, &org); err != nil {
		return nil, err
	}

	if name == "" {
		name = "ad-hoc (" + tenant + ")"
	}

	// ad-hoc with a remember request: save the profile (secret to keychain)
	if req.ProfileID == "" && req.RememberAs != "" {
		if _, serr := c.store.Save(secrets.Profile{
			Name:     req.RememberAs,
			TenantID: tenant,
			ClientID: client,
			AuthMode: mode,
		}, req.Secret); serr != nil {
			return nil, serr
		}
		name = req.RememberAs
	}

	c.s.SetClient(gc, name)
	c.s.Record("session.connect", tenant, "mode="+mode, nil)

	return &Status{
		Connected:   true,
		ProfileName: name,
		ReadOnly:    c.s.ReadOnly(),
		Org:         org,
	}, nil
}

// Domains returns the tenant's verified domain names (for UPN autocomplete).
// Returns an empty list on error so the UI degrades gracefully.
func (c *ConnectService) Domains() []string {
	client, err := c.s.Client()
	if err != nil {
		return []string{}
	}
	var resp struct {
		Value []struct {
			ID        string `json:"id"`
			IsDefault bool   `json:"isDefault"`
		} `json:"value"`
	}
	if err := client.Get(c.s.Ctx(), "/domains", nil, &resp); err != nil {
		return []string{}
	}
	// default domain first
	domains := make([]string, 0, len(resp.Value))
	for _, d := range resp.Value {
		if d.IsDefault {
			domains = append(domains, d.ID)
		}
	}
	for _, d := range resp.Value {
		if !d.IsDefault {
			domains = append(domains, d.ID)
		}
	}
	return domains
}

func (c *ConnectService) Disconnect() {
	c.s.Record("session.disconnect", c.s.ProfileName(), "", nil)
	c.s.Disconnect()
}

func (c *ConnectService) GetStatus() *Status {
	return &Status{
		Connected:   c.s.Connected(),
		ProfileName: c.s.ProfileName(),
		ReadOnly:    c.s.ReadOnly(),
	}
}

func (c *ConnectService) SetReadOnly(v bool) *Status {
	c.s.SetReadOnly(v)
	c.s.Record("session.readOnly", "", "enabled="+boolStr(v), nil)
	return c.GetStatus()
}

func boolStr(v bool) string {
	if v {
		return "true"
	}
	return "false"
}
