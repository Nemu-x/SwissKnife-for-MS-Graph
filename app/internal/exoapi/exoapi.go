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
	"math/rand/v2"
	"net/http"
	"net/url"
	"reflect"
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
	// conn is the connection's Graph client: the key under which a removed
	// or changed endpoint is remembered for the Provider.
	conn *graphapi.Client
}

// exchangeTokens adapts the session broker to graphapi.TokenSource.
type exchangeTokens struct{ b session.TokenBroker }

func (t exchangeTokens) Token(ctx context.Context) (string, error) {
	return t.b.TokenFor(ctx, auth.ResourceExchange)
}

// clientKey identifies one connection's Admin API client.
type clientKey struct {
	tokens session.TokenBroker
	base   string
}

// clients caches one client (and so one HTTP transport and its connection
// pool) per connection instead of building them for every call. Bounded:
// a long-running app reconnects, a worker serves several tenants.
var (
	clientsMu sync.Mutex
	clients   = map[clientKey]*cachedClient{}
)

type cachedClient struct {
	c    *Client
	used time.Time
}

const maxCachedClients = 16

// FromEnv returns the client for the engine's current connection.
func FromEnv(env engine.Env) (*Client, error) {
	if env.Tokens == nil || env.TenantID == "" {
		return nil, session.ErrNotConnected
	}
	base := strings.TrimRight(BaseURL, "/") + "/adminapi/v2.0/" + url.PathEscape(env.TenantID)
	// A broker that cannot be a map key (never the case for the session's
	// pointer types) simply gets a fresh client.
	if !reflect.TypeOf(env.Tokens).Comparable() {
		return newClient(env, base), nil
	}
	key := clientKey{env.Tokens, base}
	clientsMu.Lock()
	defer clientsMu.Unlock()
	if cc := clients[key]; cc != nil && cc.c.conn == env.Graph {
		cc.used = time.Now()
		return cc.c, nil
	}
	if _, replacing := clients[key]; !replacing && len(clients) >= maxCachedClients {
		// Evict only the least recently used connection's client.
		var oldest clientKey
		var at time.Time
		for k, cc := range clients {
			if at.IsZero() || cc.used.Before(at) {
				oldest, at = k, cc.used
			}
		}
		delete(clients, oldest)
	}
	c := newClient(env, base)
	clients[key] = &cachedClient{c: c, used: time.Now()}
	return c, nil
}

func newClient(env engine.Env, base string) *Client {
	return &Client{g: graphapi.New(exchangeTokens{env.Tokens}, graphapi.WithBaseURL(base)), base: base, conn: env.Graph}
}

// sleep waits between retries; tests swap it.
var sleep = func(ctx context.Context, d time.Duration) error {
	t := time.NewTimer(d)
	defer t.Stop()
	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-t.C:
		return nil
	}
}

// maxUnavailableRetries bounds the retries of a 503/504 answer.
const maxUnavailableRetries = 3

// post sends one Admin API request. Every call is a POST, which the Graph
// client does not retry on 503/504 (a POST may create something twice);
// here a POST invokes a Get- or Set- cmdlet with absolute values, which is
// safe to send again, so a busy or timed-out gateway is retried with backoff.
func (c *Client) post(ctx context.Context, path string, hdr http.Header, body any, out *json.RawMessage) error {
	backoff := time.Second
	for try := 0; ; try++ {
		err := c.g.DoWithHeaders(ctx, http.MethodPost, path, nil, hdr, body, out)
		var ge *graphapi.GraphError
		if err == nil || try >= maxUnavailableRetries || !errors.As(err, &ge) || ge.StatusCode != 503 && ge.StatusCode != 504 {
			return err
		}
		if serr := sleep(ctx, retryDelay(ge.RetryAfter, backoff)); serr != nil {
			if ctx.Err() != nil {
				return ctx.Err() // cancelled while waiting, not a gateway failure
			}
			return err
		}
		backoff *= 2
	}
}

// maxRetryWait caps a server-requested wait, as the Graph client does.
const maxRetryWait = 60 * time.Second

// retryDelay honours Retry-After when the gateway sends one, else uses the
// exponential backoff with up to 50% jitter so parallel callers spread out.
func retryDelay(retryAfter, backoff time.Duration) time.Duration {
	if retryAfter > 0 {
		return min(retryAfter, maxRetryWait)
	}
	return backoff + time.Duration(rand.Int64N(int64(backoff)/2+1))
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
	out, err := c.invoke(ctx, endpoint, cmdlet, params, anchor)
	if err != nil && endpointGone(err) {
		// The endpoint moved or was retired (the API is in Preview): stop
		// routing this connection here so the resolver picks the PowerShell
		// implementation from now on.
		markGone(c.conn)
		return nil, &EndpointChangedError{Endpoint: endpoint, Err: err}
	}
	return out, err
}

// invoke is Invoke without the endpoint-changed fallback (the availability
// probe classifies those answers itself).
func (c *Client) invoke(ctx context.Context, endpoint, cmdlet string, params map[string]any, anchor string) ([]json.RawMessage, error) {
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
		if err := c.post(ctx, next, hdr, body, &raw); err != nil {
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

// EndpointChangedError: the Admin API no longer serves an endpoint the way
// this app calls it. The connection falls back to PowerShell; the action has
// to be run again to use it. No Unwrap on purpose: the caller's Graph error
// formatting would replace this message with the bare HTTP error.
type EndpointChangedError struct {
	Endpoint string
	Err      error
}

func (e *EndpointChangedError) Error() string {
	return fmt.Sprintf("the Exchange Admin API did not accept the %s call (%v) — it may have changed while in Preview. "+
		"This connection now uses Exchange PowerShell for it: run the action again", e.Endpoint, e.Err)
}

// endpointGone tells an API-level refusal (endpoint missing, changed or not
// implemented) from a cmdlet failure: the latter carries the Exchange
// exception ("…|Microsoft.Exchange…Exception|…") and must not switch the
// backend (a mailbox that does not exist is the operator's input, not the API).
func endpointGone(err error) bool {
	var ge *graphapi.GraphError
	if !errors.As(err, &ge) {
		return false
	}
	switch ge.StatusCode {
	case 501:
		return true
	case 400, 404:
		return !strings.Contains(ge.Message, "Exception") && !strings.Contains(ge.Message, "Microsoft.Exchange")
	}
	return false
}

// gone holds the connections (their Graph client) whose Admin API answered
// with endpointGone; the Provider reports them unavailable.
var (
	goneMu sync.Mutex
	gone   = map[*graphapi.Client]time.Time{} // → when it was marked
)

// maxGone bounds the marked connections; past it only the oldest mark is
// dropped (a long-gone connection), never the others'.
const maxGone = 64

func markGone(conn *graphapi.Client) {
	if conn == nil {
		return
	}
	goneMu.Lock()
	defer goneMu.Unlock()
	if _, ok := gone[conn]; !ok && len(gone) >= maxGone {
		var oldest *graphapi.Client
		var at time.Time
		for k, t := range gone {
			if oldest == nil || t.Before(at) {
				oldest, at = k, t
			}
		}
		delete(gone, oldest)
	}
	gone[conn] = time.Now()
}

func isGone(conn *graphapi.Client) bool {
	goneMu.Lock()
	defer goneMu.Unlock()
	_, ok := gone[conn]
	return ok
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
	// A call on this connection found an endpoint gone: that overrides the
	// probe (which only tried Get-Mailbox) for the rest of the connection.
	if isGone(c) {
		p.client, p.probedAt = c, p.now()
		p.verdict = &engine.Reason{Key: "exoApiChanged"}
		return p.verdict
	}
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
	_, err = c.invoke(ctx, "Mailbox", "Get-Mailbox", map[string]any{"ResultSize": 1}, anchor)
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
