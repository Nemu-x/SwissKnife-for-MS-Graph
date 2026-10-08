package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"strings"

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
	if p.AuthMode == string(auth.ModeClientCertificate) && p.CertPath == "" {
		return secrets.Profile{}, errors.New("certificate profiles need a certificate file (.pfx)")
	}
	// Switching the auth mode or the certificate file invalidates the stored
	// secret unless a new one comes with the switch.
	if p.ID != "" && secret == "" {
		list, err := c.store.List()
		if err != nil {
			return secrets.Profile{}, err
		}
		for _, old := range list {
			// A new auth mode or a different PFX makes the old secret meaningless.
			if old.ID == p.ID && (old.AuthMode != p.AuthMode || old.CertPath != p.CertPath) {
				if err := c.store.ClearSecret(p.ID); err != nil {
					return secrets.Profile{}, err
				}
			}
		}
	}
	if p.ID != "" {
		if old, ok := c.findProfile(p.ID); ok {
			p.Policy = old.Policy
		}
	} else {
		p.Policy = nil
	}
	pending := false
	if p.AuthMode == string(auth.ModeClientCertificate) && secret == "" {
		secret, pending = c.store.PendingCertPassword(p.CertPath)
	}
	saved, err := c.store.Save(p, secret)
	if err == nil && pending {
		c.store.DropPendingCertPassword(p.CertPath)
	}
	return saved, err
}

func (c *ConnectService) findProfile(id string) (secrets.Profile, bool) {
	list, err := c.store.List()
	if err != nil {
		return secrets.Profile{}, false
	}
	for _, p := range list {
		if p.ID == id {
			return p, true
		}
	}
	return secrets.Profile{}, false
}

// SetProfilePolicy changes a profile's policy; the connected profile gets it
// at once. maxDanger is "" (no limit), read, write or destructive;
// allowedGroups are group ids (empty = any target).
func (c *ConnectService) SetProfilePolicy(profileID string, p session.Policy) (secrets.Profile, error) {
	switch p.MaxDanger {
	case "", "read", "write", "destructive":
	default:
		return secrets.Profile{}, errors.New("maxDanger must be read, write or destructive")
	}
	prof, ok := c.findProfile(profileID)
	if !ok {
		return secrets.Profile{}, errors.New("profile not found")
	}
	// Limits cannot be lifted from inside the session they limit.
	live := c.s.Connected() && c.s.ProfileID() == prof.ID
	if live {
		if cur := c.s.Policy(); cur.MaxDanger != "" || cur.Scoped() {
			return secrets.Profile{}, &OpError{Code: "policyLive", Message: "disconnect before changing the limits of the profile you are connected with"}
		}
	}
	clean := []string{}
	seen := map[string]bool{}
	for _, g := range p.AllowedGroups {
		g = strings.TrimSpace(g)
		if g != "" && !seen[g] {
			seen[g] = true
			clean = append(clean, g)
		}
	}
	p.AllowedGroups = clean
	labels := map[string]string{}
	for _, g := range clean {
		if l := strings.TrimSpace(p.GroupLabels[g]); l != "" {
			labels[g] = l
		}
	}
	p.GroupLabels = labels
	if p.MaxDanger == "" && len(clean) == 0 {
		prof.Policy = nil
	} else {
		prof.Policy = &p
	}
	saved, err := c.store.Save(prof, "")
	if err != nil {
		return secrets.Profile{}, err
	}
	if live {
		c.s.SetPolicy(prof.ID, p)
	}
	c.s.Record("profile.policy", prof.Name, fmt.Sprintf("maxDanger=%s allowedGroups=%d", p.MaxDanger, len(clean)), nil)
	return saved, nil
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
	AuthMode   string `json:"authMode"` // client_secret | device_code | client_certificate
	CertPath   string `json:"certPath"` // client_certificate: PFX file; Secret is its password
	RememberAs string `json:"rememberAs"`
}

// Status is the session state exposed to the frontend.
type Status struct {
	Connected   bool            `json:"connected"`
	ProfileName string          `json:"profileName"`
	ReadOnly    bool            `json:"readOnly"`
	Policy      session.Policy  `json:"policy"`
	ProfileID   string          `json:"profileId,omitempty"`
	Org         json.RawMessage `json:"org,omitempty"`
}

// Credentials is a resolved set of connection parameters. The secret lives in
// memory only for the lifetime of the connect call (ADR-002).
type Credentials struct {
	TenantID string
	ClientID string
	Secret   string
	AuthMode string // client_secret | device_code | client_certificate
	CertPath string // client_certificate: PFX file; Secret is its password
	Name     string // profile name; empty for ad-hoc credentials
	// ProfileID and Policy come from a saved profile (empty when ad hoc).
	ProfileID string
	Policy    session.Policy
	// DelegatedOrg: a partner profile signs in to this customer tenant.
	DelegatedOrg string
}

// ResolveCredentials turns a ConnectRequest into concrete credentials. A
// ProfileID loads the saved profile and, for app-only mode, its keychain
// secret; otherwise the ad-hoc fields are used as given. Shared by the GUI
// ConnectService and the headless CLI so both connect the same way.
func ResolveCredentials(store *secrets.Store, req ConnectRequest) (Credentials, error) {
	cr := Credentials{
		TenantID: req.TenantID,
		ClientID: req.ClientID,
		Secret:   req.Secret,
		AuthMode: req.AuthMode,
		CertPath: req.CertPath,
	}
	if req.ProfileID == "" {
		if cr.AuthMode == string(auth.ModeClientCertificate) && cr.Secret == "" {
			cr.Secret, _ = store.PendingCertPassword(cr.CertPath)
		}
		return cr, nil
	}
	profiles, err := store.List()
	if err != nil {
		return Credentials{}, err
	}
	var found *secrets.Profile
	for i := range profiles {
		if profiles[i].ID == req.ProfileID {
			found = &profiles[i]
			break
		}
	}
	if found == nil {
		return Credentials{}, errors.New("profile not found")
	}
	cr.TenantID, cr.ClientID, cr.AuthMode, cr.Name = found.TenantID, found.ClientID, found.AuthMode, found.Name
	cr.CertPath, cr.ProfileID = found.CertPath, found.ID
	// GDAP: the partner's user signs in to the customer tenant directly.
	if found.DelegatedOrg != "" && cr.AuthMode == string(auth.ModeDeviceCode) {
		cr.TenantID, cr.DelegatedOrg = found.DelegatedOrg, found.DelegatedOrg
	}
	if found.Policy != nil {
		cr.Policy = *found.Policy
	}
	switch cr.AuthMode {
	case string(auth.ModeClientSecret):
		cr.Secret, err = store.Secret(found.ID)
		if err != nil {
			return Credentials{}, err
		}
	case string(auth.ModeClientCertificate):
		// A PFX without a password is valid; a missing keychain entry is "".
		if found.HasSecret {
			cr.Secret, err = store.Secret(found.ID)
			if err != nil {
				return Credentials{}, err
			}
		}
	}
	return cr, nil
}

// NewTokenProvider builds the Graph token source for the credentials. prompt
// is invoked for device-code sign-in (the caller decides how to show the code:
// a Wails event in the GUI, stderr in the CLI); it is ignored for app-only.
func NewTokenProvider(cr Credentials, prompt auth.DeviceCodePrompt) (*auth.TokenProvider, error) {
	switch cr.AuthMode {
	case string(auth.ModeDeviceCode):
		return auth.NewDeviceCode(cr.TenantID, cr.ClientID, prompt)
	case string(auth.ModeClientCertificate):
		if cr.CertPath == "" {
			return nil, errors.New("choose or generate a certificate file (.pfx) first")
		}
		pfx, err := os.ReadFile(cr.CertPath)
		if err != nil {
			return nil, fmt.Errorf("certificate file: %w", err)
		}
		return auth.NewClientCertificate(cr.TenantID, cr.ClientID, pfx, cr.Secret)
	default:
		if cr.Secret == "" {
			return nil, errors.New("client secret is required")
		}
		return auth.NewClientSecret(cr.TenantID, cr.ClientID, cr.Secret)
	}
}

// Connect establishes the session and runs a self-test GET /organization.
func (c *ConnectService) Connect(req ConnectRequest) (*Status, error) {
	cr, err := ResolveCredentials(c.store, req)
	if err != nil {
		return nil, err
	}
	tenant, client, mode, name := cr.TenantID, cr.ClientID, cr.AuthMode, cr.Name

	provider, err := NewTokenProvider(cr, func(url, code, msg string) {
		wrt.EventsEmit(c.s.Ctx(), "auth:deviceCode", map[string]string{
			"url": url, "code": code, "message": msg,
		})
	})
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
			CertPath: cr.CertPath,
		}, cr.Secret); serr != nil {
			return nil, serr
		}
		if mode == string(auth.ModeClientCertificate) {
			c.store.DropPendingCertPassword(cr.CertPath)
		}
		name = req.RememberAs
	}

	// The policy goes in before the client, so no write runs unlimited.
	c.s.SetPolicy(cr.ProfileID, cr.Policy)
	c.s.SetClient(gc, name)
	c.s.SetTokens(provider)
	c.s.SetIdentity(tenant, mode != string(auth.ModeDeviceCode))
	c.s.SetDelegatedOrg(cr.DelegatedOrg)
	c.s.Record("session.connect", tenant, "mode="+mode, nil)

	st := c.GetStatus()
	st.Org = org
	return st, nil
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
		Policy:      c.s.Policy(),
		ProfileID:   c.s.ProfileID(),
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
