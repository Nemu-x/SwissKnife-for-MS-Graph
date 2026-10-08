package pwsh

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"net/url"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
)

// Module families and the backends they serve.
const (
	FamilyExchange = "exo"
	FamilyTeams    = "teams"

	BackendExchangePS engine.Backend = "exo-ps"
	BackendTeamsPS    engine.Backend = "teams-ps"
)

// Connect-* takes a fixed access token and cannot refresh it, so the host is
// signed in again shortly before that token expires.
const (
	renewBefore    = 5 * time.Minute
	fallbackTTL    = 45 * time.Minute // when the token's exp cannot be read
	connectTimeout = 2 * time.Minute
)

type entry struct {
	host       *Host
	conn       *graphapi.Client // connection the host is signed in for
	validUntil time.Time        // zero = not signed in
}

// Pool keeps one signed-in host per module family for the current connection
// and implements engine.PSRunner.
type Pool struct {
	det   *Detector
	allow map[string][]string // family → cmdlets the host may run

	mu    sync.Mutex
	hosts map[string]*entry
	gen    uint64                    // bumped by Close: a host started before it is discarded
	closed map[*graphapi.Client]bool // connections whose hosts were closed
	start func(exe string, allow []string) (*Host, error)
	now   func() time.Time
}

func NewPool(det *Detector, allow map[string][]string) *Pool {
	return &Pool{det: det, allow: allow, hosts: map[string]*entry{}, closed: map[*graphapi.Client]bool{}, start: Start, now: time.Now}
}

// Invoke implements engine.PSRunner. An authentication failure (the token
// was revoked or expired early) signs the host in again; a read (Get-*) is
// then retried once, a write is not — repeating it could repeat a change.
func (p *Pool) Invoke(env engine.Env, family, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error) {
	for attempt := 0; ; attempt++ {
		h, err := p.host(env, family)
		if err != nil {
			return nil, err
		}
		out, err := h.Invoke(env.Ctx, cmdlet, params, sel...)
		if errors.Is(err, ErrHostExited) || !h.Alive() {
			p.drop(family, h)
		}
		if isAuthError(err) {
			p.expire(family, h)
			if attempt == 0 && strings.HasPrefix(cmdlet, "Get-") {
				continue
			}
		}
		return out, err
	}
}

func isAuthError(err error) bool {
	var pe *Error
	if !errors.As(err, &pe) {
		return false
	}
	if strings.EqualFold(pe.Category, "AuthenticationError") {
		return true
	}
	m := strings.ToLower(pe.Message)
	for _, phrase := range []string{"401 unauthorized", "(401) unauthorized", "token has expired", "token is expired", "access token expired", "lifetime validation failed"} {
		if strings.Contains(m, phrase) {
			return true
		}
	}
	return false
}

// expire forces the next call to sign the host in again.
func (p *Pool) expire(family string, h *Host) {
	p.mu.Lock()
	defer p.mu.Unlock()
	if e := p.hosts[family]; e != nil && e.host == h {
		e.validUntil = time.Time{}
	}
}

// host returns a connected host, (re)starting or reconnecting as needed.
func (p *Pool) host(env engine.Env, family string) (*Host, error) {
	p.mu.Lock()
	defer p.mu.Unlock()
	// A call from a connection that already ended must not touch the pool —
	// the family's host may belong to a newer connection by now.
	if p.closed[env.Graph] {
		return nil, errors.New("disconnected")
	}
	e := p.hosts[family]
	if e != nil && (!e.host.Alive() || e.conn != env.Graph) {
		e.host.Close()
		delete(p.hosts, family)
		e = nil
	}
	if e == nil {
		pe := p.det.Get(env.Ctx)
		if !pe.PwshOK() {
			return nil, errors.New("PowerShell 7.2 or later is not installed")
		}
		gen := p.gen
		h, err := p.start(pe.Exe, p.allow[family])
		if err != nil {
			return nil, err
		}
		if gen != p.gen || p.closed[env.Graph] {
			h.Close()
			return nil, errors.New("disconnected")
		}
		e = &entry{host: h, conn: env.Graph}
		p.hosts[family] = e
	}
	if e.validUntil.IsZero() || p.now().After(e.validUntil) {
		params, err := connectParams(env, family)
		if err != nil {
			return nil, err
		}
		ctx, cancel := context.WithTimeout(env.Ctx, connectTimeout)
		err = e.host.Connect(ctx, family, params)
		cancel()
		if err != nil {
			if !e.host.Alive() {
				delete(p.hosts, family)
			}
			return nil, err
		}
		e.validUntil = signInValidUntil(params["token"], p.now())
	}
	return e.host, nil
}

// signInValidUntil reads the token's exp claim (no signature check: the
// token is ours, only its lifetime matters) and backs off renewBefore.
func signInValidUntil(token any, now time.Time) time.Time {
	tok, _ := token.(string)
	parts := strings.Split(tok, ".")
	if len(parts) == 3 {
		if raw, err := base64.RawURLEncoding.DecodeString(parts[1]); err == nil {
			var claims struct {
				Exp int64 `json:"exp"`
			}
			if json.Unmarshal(raw, &claims) == nil && claims.Exp > 0 {
				return time.Unix(claims.Exp, 0).Add(-renewBefore)
			}
		}
	}
	return now.Add(fallbackTTL)
}

func (p *Pool) drop(family string, h *Host) {
	p.mu.Lock()
	defer p.mu.Unlock()
	if e := p.hosts[family]; e != nil && e.host == h {
		h.Close()
		delete(p.hosts, family)
	}
}

// CloseFor stops the hosts signed in for one connection (it ended). Hosts of
// a newer connection stay; a start still in flight for conn is discarded.
func (p *Pool) CloseFor(conn *graphapi.Client) {
	if conn == nil {
		return
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	p.closed[conn] = true
	for f, e := range p.hosts {
		if e.conn == conn {
			e.host.Close()
			delete(p.hosts, f)
		}
	}
}

// Close stops every host (shutdown).
func (p *Pool) Close() {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.gen++
	for f, e := range p.hosts {
		e.host.Close()
		delete(p.hosts, f)
	}
}

// connectParams builds the Connect-* arguments: tokens from the broker, plus
// the organization (app-only; Exchange wants the initial .onmicrosoft.com
// domain) or the signed-in admin's UPN (delegated).
func connectParams(env engine.Env, family string) (map[string]any, error) {
	if env.Tokens == nil || env.Graph == nil {
		return nil, errors.New("not connected")
	}
	switch family {
	case FamilyTeams:
		graphTok, err := env.Tokens.TokenFor(env.Ctx, auth.ResourceGraph)
		if err != nil {
			return nil, err
		}
		teamsTok, err := env.Tokens.TokenFor(env.Ctx, auth.ResourceTeams)
		if err != nil {
			return nil, err
		}
		return map[string]any{"graphToken": graphTok, "token": teamsTok}, nil
	default:
		tok, err := env.Tokens.TokenFor(env.Ctx, auth.ResourceExchange)
		if err != nil {
			return nil, err
		}
		params := map[string]any{"token": tok}
		switch {
		case env.AppOnly:
			org, err := graphapi.InitialDomain(env.Ctx, env.Graph)
			if err != nil {
				return nil, err
			}
			params["organization"] = org
		case env.DelegatedOrg != "":
			// GDAP: the partner's user is no object of the customer
			// directory, so there is no /me; the customer names the org.
			org, err := graphapi.InitialDomain(env.Ctx, env.Graph)
			if err != nil {
				return nil, err
			}
			params["organization"], params["delegatedOrg"] = org, org
		default:
			var me struct {
				UPN string `json:"userPrincipalName"`
			}
			if err := env.Graph.Get(env.Ctx, "/me", url.Values{"$select": {"userPrincipalName"}}, &me); err != nil {
				return nil, err
			}
			params["upn"] = me.UPN
		}
		return params, nil
	}
}
