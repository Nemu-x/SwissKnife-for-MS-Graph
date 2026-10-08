// Package exoapi talks to the Exchange Online Admin API (ADR-008): a REST
// wrapper around selected Exchange cmdlets, called with an Exchange token
// instead of PowerShell. It reuses the Graph client for transport, retries and
// 429 handling; only the base URL, the token resource and the per-call
// X-AnchorMailbox routing hint differ.
//
// The API is in Preview and covers six endpoints; every action that uses it
// also declares a PowerShell implementation the resolver falls back to.
package exoapi

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/session"
)

// BackendExoAPI is the engine backend served by this package.
const BackendExoAPI engine.Backend = "exo-api"

// BaseURL is the Microsoft 365 / GCC host; tests override it.
var BaseURL = "https://outlook.office365.com"

// systemMailbox is the routing key for app-only calls that target no
// particular mailbox; the GUID is the same in every organization.
const systemMailbox = "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}"

// Client calls one tenant's Admin API.
type Client struct {
	g    *graphapi.Client
	base string // https://<host>/adminapi/v2.0/<tenant>
}

// exchangeTokens adapts the session broker to graphapi.TokenSource.
type exchangeTokens struct{ b session.TokenBroker }

func (t exchangeTokens) Token(ctx context.Context) (string, error) {
	return t.b.TokenFor(ctx, auth.ResourceExchange)
}

// FromEnv builds a client for the engine's current connection.
func FromEnv(env engine.Env) (*Client, error) {
	if env.Tokens == nil || env.TenantID == "" {
		return nil, session.ErrNotConnected
	}
	base := strings.TrimRight(BaseURL, "/") + "/adminapi/v2.0/" + url.PathEscape(env.TenantID)
	return &Client{g: graphapi.New(exchangeTokens{env.Tokens}, graphapi.WithBaseURL(base)), base: base}, nil
}

// MailboxAnchor is the routing hint for a call bound to one mailbox.
func MailboxAnchor(mailbox string) string { return "UPN:" + mailbox }

// OrgAnchor is the routing hint for org-wide app-only calls; domain is the
// tenant's initial .onmicrosoft.com domain (graphapi.InitialDomain).
func OrgAnchor(domain string) string { return "APP:" + systemMailbox + "@" + domain }

// Invoke runs one cmdlet on an endpoint and returns the result objects,
// following @odata.nextLink pages (re-POSTing the same body, as the API
// requires). Writes return no objects.
func (c *Client) Invoke(ctx context.Context, endpoint, cmdlet string, params map[string]any, anchor string) ([]json.RawMessage, error) {
	input := map[string]any{"CmdletName": cmdlet}
	if len(params) > 0 {
		input["Parameters"] = params
	}
	body := map[string]any{"CmdletInput": input}
	hdr := http.Header{"X-Anchormailbox": {anchor}}

	var out []json.RawMessage
	next := "/" + endpoint
	seen := map[string]bool{}
	for next != "" {
		// A service that hands back the same continuation would loop forever.
		if seen[next] {
			return nil, errors.New("exchange admin api: pagination did not advance")
		}
		seen[next] = true
		// The bearer token must never follow a link to another host.
		if strings.Contains(next, "://") && !strings.HasPrefix(next, c.base+"/") {
			return nil, errors.New("exchange admin api: continuation link points to another host")
		}
		var raw json.RawMessage
		if err := c.g.DoWithHeaders(ctx, http.MethodPost, next, nil, hdr, body, &raw); err != nil {
			return nil, err
		}
		items, link, err := decode(raw)
		if err != nil {
			return nil, err
		}
		out = append(out, items...)
		next = link
	}
	return out, nil
}

// decode accepts the shapes the API answers with: an OData page
// ({"value": [...]}), a bare array, a single object, or nothing.
func decode(raw json.RawMessage) ([]json.RawMessage, string, error) {
	trimmed := strings.TrimSpace(string(raw))
	switch {
	case trimmed == "" || trimmed == "null":
		return nil, "", nil
	case strings.HasPrefix(trimmed, "["):
		var arr []json.RawMessage
		if err := json.Unmarshal(raw, &arr); err != nil {
			return nil, "", fmt.Errorf("exchange admin api: decode response: %w", err)
		}
		return arr, "", nil
	}
	var page struct {
		Value    *[]json.RawMessage `json:"value"`
		NextLink string             `json:"@odata.nextLink"`
	}
	if json.Unmarshal(raw, &page) == nil && page.Value != nil {
		return *page.Value, page.NextLink, nil
	}
	return []json.RawMessage{raw}, "", nil
}

// Provider reports whether the Admin API answers for the connected tenant.
// The API is in Preview and not enabled everywhere, so the first status
// check after a connect probes it once (Get-Mailbox for one mailbox); the
// verdict is cached per connection.
type Provider struct {
	mu       sync.Mutex
	client   *graphapi.Client // connection the cached verdict belongs to
	verdict  *engine.Reason
	probedAt time.Time
	now      func() time.Time
	// probe is swapped in tests.
	probe func(ctx context.Context, env engine.Env) error
}

func NewProvider() *Provider { return &Provider{probe: defaultProbe, now: time.Now} }

// retryTransient is how long a transient probe failure (network, throttling,
// token) is trusted before the next status check probes again; definitive
// answers (works, permission, not enabled) hold for the whole connection.
const retryTransient = time.Minute

func (p *Provider) Backend() engine.Backend { return BackendExoAPI }

func (p *Provider) Status(s *session.Session) *engine.Reason {
	c, err := s.Client()
	if err != nil || s.Tokens() == nil {
		return &engine.Reason{Key: "notConnected"}
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	if p.client == c && !p.probedAt.IsZero() &&
		(p.verdict == nil || p.verdict.Key != "exoUnreachable" || p.now().Sub(p.probedAt) < retryTransient) {
		return p.verdict
	}
	ctx, cancel := context.WithTimeout(s.Ctx(), 10*time.Second)
	defer cancel()
	env := engine.Env{Ctx: ctx, Graph: c, Tokens: s.Tokens(), TenantID: s.TenantID(), AppOnly: s.AppOnly()}
	p.verdict = classify(p.probe(ctx, env))
	p.client, p.probedAt = c, p.now()
	return p.verdict
}

func defaultProbe(ctx context.Context, env engine.Env) error {
	c, err := FromEnv(env)
	if err != nil {
		return err
	}
	// The routing key comes from Graph; a failure there is not an Exchange
	// verdict, so it is reported as transient.
	var anchor string
	if env.AppOnly || env.DelegatedOrg != "" { // GDAP partners have no mailbox there
		domain, err := graphapi.InitialDomain(ctx, env.Graph)
		if err != nil {
			return transient{err}
		}
		anchor = OrgAnchor(domain)
	} else {
		// Delegated calls route on the signed-in admin's mailbox.
		var me struct {
			UPN string `json:"userPrincipalName"`
		}
		if err := env.Graph.Get(ctx, "/me", url.Values{"$select": {"userPrincipalName"}}, &me); err != nil {
			return transient{err}
		}
		anchor = MailboxAnchor(me.UPN)
	}
	_, err = c.Invoke(ctx, "Mailbox", "Get-Mailbox", map[string]any{"ResultSize": 1}, anchor)
	return err
}

// transient marks a probe failure outside Exchange itself.
type transient struct{ error }

func (t transient) Unwrap() error { return t.error }

// classify turns the probe outcome into availability: a permission problem
// and a tenant without the Preview are different fixes for the operator.
func classify(err error) *engine.Reason {
	if err == nil {
		return nil
	}
	var tr transient
	if errors.As(err, &tr) {
		return &engine.Reason{Key: "exoUnreachable"}
	}
	var ge *graphapi.GraphError
	if errors.As(err, &ge) {
		switch ge.StatusCode {
		case 401, 403:
			return &engine.Reason{Key: "exoPermission"}
		case 400, 404:
			return &engine.Reason{Key: "exoApiNotEnabled"}
		}
	}
	return &engine.Reason{Key: "exoUnreachable"}
}
