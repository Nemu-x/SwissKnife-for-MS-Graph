package worker

import (
	"context"
	"crypto/tls"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
)

// InvokeRequest is one cmdlet call for a client's tenant.
type InvokeRequest struct {
	Family       string         `json:"family"`
	Cmdlet       string         `json:"cmdlet"`
	Params       map[string]any `json:"params,omitempty"`
	Select       []string       `json:"select,omitempty"`
	TenantID     string         `json:"tenantId"`
	AppOnly      bool           `json:"appOnly"`
	DelegatedOrg string         `json:"delegatedOrg,omitempty"`
	// ExchangeOrg / ExchangeUPN let Exchange sign in without a Graph token.
	ExchangeOrg string            `json:"exchangeOrg,omitempty"`
	ExchangeUPN string            `json:"exchangeUpn,omitempty"`
	Tokens      map[string]string `json:"tokens"` // resource → access token
}

// InvokeResponse carries the objects or the PowerShell error.
type InvokeResponse struct {
	Data  []json.RawMessage `json:"data"`
	Error *pwsh.Error       `json:"error,omitempty"`
}

// Health is what a paired client learns about the worker.
type Health struct {
	Name     string   `json:"name"`
	Version  string   `json:"version"`
	Families []string `json:"families"` // module families it will run
	// Unavailable are families the operator enabled whose module is not
	// usable there now (PowerShell or the module missing).
	Unavailable []string `json:"unavailable,omitempty"`
}

// PairedClient is a client allowed to call the worker.
type PairedClient struct {
	Fingerprint string    `json:"fingerprint"`
	Name        string    `json:"name"`
	PairedAt    time.Time `json:"pairedAt"`
}

// Server is the worker side.
type Server struct {
	Identity *Identity
	Dir      string // clients.json lives here
	Name     string
	Version  string
	// NewRunner makes the PowerShell runner of one client connection: each
	// gets its own hosts, so clients and tenants never share (or kill) a
	// signed-in process.
	NewRunner func() engine.PSRunner
	// Families the operator enabled; Available narrows them to the modules
	// actually installed (nil = all enabled ones).
	Families  map[string]bool
	Available func() map[string]bool
	Allow     map[string][]string // family → allow-listed cmdlets
	Audit     *auditlog.Log
	Logf      func(format string, a ...any)

	mu        sync.Mutex
	clients   map[string]PairedClient
	clientsAt time.Time // modification time of the loaded clients file
	pair      *pairing
	conns     map[string]*conn
}

// Limits on what one paired client can make the worker hold.
const (
	maxConnsPerClient = 4
	maxConns          = 16
)

// conn is one client's tenant connection: a stable Graph client (the pool
// keys signed-in hosts by it) whose tokens each call refreshes.
type conn struct {
	client string // client fingerprint
	graph  *graphapi.Client
	toks   *tokens
	runner engine.PSRunner
	used   time.Time
}

// close stops the connection's hosts and forgets its tokens.
func (c *conn) close() {
	c.toks.set(nil)
	if cl, ok := c.runner.(interface{ Close() }); ok {
		cl.Close()
	}
}

type tokens struct {
	mu sync.Mutex
	m  map[string]string
}

func (t *tokens) set(m map[string]string) {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.m = m
}

func (t *tokens) TokenFor(_ context.Context, resource string) (string, error) {
	t.mu.Lock()
	defer t.mu.Unlock()
	if v := t.m[resource]; v != "" {
		return v, nil
	}
	return "", fmt.Errorf("the client sent no token for %s", resource)
}

func (t *tokens) Token(ctx context.Context) (string, error) {
	return t.TokenFor(ctx, auth.ResourceGraph)
}

func (s *Server) clientsPath() string { return filepath.Join(s.Dir, "worker-clients.json") }

// Clients lists the paired clients.
func (s *Server) Clients() []PairedClient {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.loadLocked()
	out := make([]PairedClient, 0, len(s.clients))
	for _, c := range s.clients {
		out = append(out, c)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].PairedAt.Before(out[j].PairedAt) })
	return out
}

// Revoke unpairs a client by fingerprint (a prefix of 8+ characters works).
func (s *Server) Revoke(fp string) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.loadLocked()
	fp = strings.ToLower(fp)
	var hit []string
	for k := range s.clients {
		if len(fp) >= 8 && strings.HasPrefix(k, fp) {
			hit = append(hit, k)
		}
	}
	if len(hit) != 1 {
		return errors.New("no single paired client has that fingerprint")
	}
	delete(s.clients, hit[0])
	s.dropUnpairedLocked(nil)
	if err := s.saveLocked(); err != nil {
		return err
	}
	s.clientsAt = time.Time{} // re-read our own write next time
	return nil
}

// loadLocked (re)reads the clients file when it changed on disk — a
// "worker revoke" from another process takes effect at the next call.
func (s *Server) loadLocked() {
	st, err := os.Stat(s.clientsPath())
	if s.clients != nil && (err != nil && s.clientsAt.IsZero() || err == nil && st.ModTime().Equal(s.clientsAt)) {
		return
	}
	prev := s.clients
	s.clients = map[string]PairedClient{}
	s.clientsAt = time.Time{}
	if err == nil {
		s.clientsAt = st.ModTime()
	}
	defer s.dropUnpairedLocked(prev)
	if b, err := os.ReadFile(s.clientsPath()); err == nil {
		var list []PairedClient
		if json.Unmarshal(b, &list) == nil {
			for _, c := range list {
				s.clients[c.Fingerprint] = c
			}
		}
	}
}

// dropUnpairedLocked closes the connections of clients no longer paired.
func (s *Server) dropUnpairedLocked(prev map[string]PairedClient) {
	for k, c := range s.conns {
		if _, ok := s.clients[c.client]; !ok {
			c.close()
			delete(s.conns, k)
		}
	}
	_ = prev
}

func (s *Server) saveLocked() error {
	list := make([]PairedClient, 0, len(s.clients))
	for _, c := range s.clients {
		list = append(list, c)
	}
	b, _ := json.MarshalIndent(list, "", "  ")
	tmp := s.clientsPath() + ".tmp"
	if err := os.WriteFile(tmp, b, 0o600); err != nil {
		return err
	}
	if err := os.Rename(tmp, s.clientsPath()); err != nil {
		return err
	}
	if st, err := os.Stat(s.clientsPath()); err == nil {
		s.clientsAt = st.ModTime()
	}
	return nil
}

func (s *Server) isPaired(fp string) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.loadLocked()
	_, ok := s.clients[fp]
	return ok
}

// StartPairing opens pairing for one client and returns the code to show.
func (s *Server) StartPairing() string {
	code := NewCode()
	s.mu.Lock()
	s.pair = newPairing(code)
	s.mu.Unlock()
	return code
}

func (s *Server) pairingActive() bool {
	s.mu.Lock()
	p := s.pair
	s.mu.Unlock()
	return p.active()
}

// TLSConfig is TLS 1.3 with a client certificate always required; outside
// pairing an unpaired certificate fails the handshake itself.
func (s *Server) TLSConfig() *tls.Config {
	// No session tickets: a resumed session skips certificate verification.
	base := &tls.Config{Certificates: []tls.Certificate{s.Identity.Cert}, MinVersion: tls.VersionTLS13,
		ClientAuth: tls.RequireAnyClientCert, SessionTicketsDisabled: true}
	return &tls.Config{
		MinVersion:             tls.VersionTLS13,
		SessionTicketsDisabled: true,
		GetConfigForClient: func(*tls.ClientHelloInfo) (*tls.Config, error) {
			c := base.Clone()
			if !s.pairingActive() {
				c.VerifyPeerCertificate = pinned(s.isPaired)
			}
			return c, nil
		},
	}
}

func peerFP(r *http.Request) string {
	if r.TLS == nil || len(r.TLS.PeerCertificates) == 0 {
		return ""
	}
	return Fingerprint(r.TLS.PeerCertificates[0].Raw)
}

// Handler serves /pair, /health and /invoke.
func (s *Server) Handler() http.Handler {
	mux := http.NewServeMux()
	mux.HandleFunc("POST /pair", s.handlePair)
	mux.HandleFunc("GET /health", s.paired(s.handleHealth))
	mux.HandleFunc("POST /invoke", s.paired(s.handleInvoke))
	mux.HandleFunc("POST /unpair", s.paired(s.handleUnpair))
	return mux
}

// paired admits only pinned clients (during pairing the handshake lets
// other certificates through, so the check is repeated here).
func (s *Server) paired(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		if fp := peerFP(r); fp == "" || !s.isPaired(fp) {
			http.Error(w, "not paired", http.StatusForbidden)
			return
		}
		next(w, r)
	}
}

func (s *Server) handlePair(w http.ResponseWriter, r *http.Request) {
	var req struct {
		Name  string `json:"name"`
		Proof string `json:"proof"`
	}
	if err := json.NewDecoder(io.LimitReader(r.Body, 4096)).Decode(&req); err != nil {
		http.Error(w, "bad request", http.StatusBadRequest)
		return
	}
	fp := peerFP(r)
	s.mu.Lock()
	p := s.pair
	s.mu.Unlock()
	if fp == "" || !p.active() {
		http.Error(w, "pairing is not open — start the worker with --pair", http.StatusForbidden)
		return
	}
	ip, _, _ := net.SplitHostPort(r.RemoteAddr)
	answer, ok := p.check(ip, fp, s.Identity.Fingerprint, req.Proof)
	if !ok {
		s.log("pairing attempt refused from %s", r.RemoteAddr)
		http.Error(w, "wrong code", http.StatusForbidden)
		return
	}
	name := strings.TrimSpace(req.Name)
	if len(name) > 80 {
		name = name[:80]
	}
	s.mu.Lock()
	s.loadLocked()
	s.clients[fp] = PairedClient{Fingerprint: fp, Name: name, PairedAt: time.Now()}
	err := s.saveLocked()
	s.mu.Unlock()
	if err != nil {
		http.Error(w, "cannot save the pairing", http.StatusInternalServerError)
		return
	}
	s.audit("worker.pair", name, "client="+fp[:16], nil)
	s.log("paired with %q (%s)", name, fp[:16])
	_ = json.NewEncoder(w).Encode(map[string]string{"proof": answer, "name": s.Name})
}

// handleUnpair lets a client remove itself.
func (s *Server) handleUnpair(w http.ResponseWriter, r *http.Request) {
	fp := peerFP(r)
	s.mu.Lock()
	s.loadLocked()
	name := s.clients[fp].Name
	delete(s.clients, fp)
	s.dropUnpairedLocked(nil)
	err := s.saveLocked()
	s.mu.Unlock()
	s.audit("worker.unpair", name, "client="+fp[:16], err)
	if err != nil {
		http.Error(w, "cannot save", http.StatusInternalServerError)
		return
	}
	w.WriteHeader(http.StatusNoContent)
}

func (s *Server) families() []string {
	out, _ := s.familyState()
	return out
}

// familyState splits the enabled families into those that run now and
// those whose module is missing.
func (s *Server) familyState() (ready, missing []string) {
	var avail map[string]bool
	if s.Available != nil {
		avail = s.Available()
	}
	for f, on := range s.Families {
		switch {
		case !on:
		case avail == nil || avail[f]:
			ready = append(ready, f)
		default:
			missing = append(missing, f)
		}
	}
	sort.Strings(ready)
	sort.Strings(missing)
	return ready, missing
}

func (s *Server) handleHealth(w http.ResponseWriter, _ *http.Request) {
	ready, missing := s.familyState()
	_ = json.NewEncoder(w).Encode(Health{Name: s.Name, Version: s.Version, Families: ready, Unavailable: missing})
}

func (s *Server) allowed(family, cmdlet string) bool {
	if !s.Families[family] {
		return false
	}
	for _, c := range s.Allow[family] {
		if strings.EqualFold(c, cmdlet) {
			return true
		}
	}
	return false
}

func (s *Server) handleInvoke(w http.ResponseWriter, r *http.Request) {
	var req InvokeRequest
	if err := json.NewDecoder(io.LimitReader(r.Body, 1<<20)).Decode(&req); err != nil {
		http.Error(w, "bad request", http.StatusBadRequest)
		return
	}
	fp := peerFP(r)
	reply := func(resp InvokeResponse) { _ = json.NewEncoder(w).Encode(resp) }
	if !s.allowed(req.Family, req.Cmdlet) {
		err := &pwsh.Error{Message: fmt.Sprintf("the worker does not run %s (%s)", req.Cmdlet, req.Family), Type: "WorkerRefused"}
		s.audit("worker.invoke", req.Cmdlet, "client="+fp[:16]+" tenant="+req.TenantID, err)
		reply(InvokeResponse{Error: err})
		return
	}
	c := s.conn(fp, req)
	env := engine.Env{Ctx: r.Context(), Graph: c.graph, Tokens: c.toks, TenantID: req.TenantID, AppOnly: req.AppOnly,
		DelegatedOrg: req.DelegatedOrg, ExchangeOrg: req.ExchangeOrg, ExchangeUPN: req.ExchangeUPN}
	data, err := c.runner.Invoke(env, req.Family, req.Cmdlet, req.Params, req.Select...)
	s.audit("worker.invoke", req.Cmdlet, "client="+fp[:16]+" tenant="+req.TenantID, err)
	if err != nil {
		var pe *pwsh.Error
		if !errors.As(err, &pe) {
			pe = &pwsh.Error{Message: err.Error(), Type: "WorkerError"}
		}
		reply(InvokeResponse{Error: pe})
		return
	}
	if data == nil {
		data = []json.RawMessage{}
	}
	reply(InvokeResponse{Data: data})
}

// conn returns the client's connection for the tenant with fresh tokens.
func (s *Server) conn(fp string, req InvokeRequest) *conn {
	key := fmt.Sprintf("%s|%s|%v|%s", fp, req.TenantID, req.AppOnly, req.DelegatedOrg)
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.conns == nil {
		s.conns = map[string]*conn{}
	}
	c := s.conns[key]
	if c == nil {
		s.evictLocked(fp)
		t := &tokens{}
		c = &conn{client: fp, graph: graphapi.New(t), toks: t, runner: s.NewRunner()}
		s.conns[key] = c
	}
	c.toks.set(req.Tokens)
	c.used = time.Now()
	return c
}

// evictLocked makes room for a new connection of client fp: the least
// recently used one goes when the client or the worker is at its limit.
func (s *Server) evictLocked(fp string) {
	for {
		mine, all := 0, len(s.conns)
		var oldestKey string
		var oldest time.Time
		for k, c := range s.conns {
			if c.client == fp {
				mine++
			}
			if (all >= maxConns || c.client == fp) && (oldestKey == "" || c.used.Before(oldest)) {
				oldestKey, oldest = k, c.used
			}
		}
		if mine < maxConnsPerClient && all < maxConns || oldestKey == "" {
			return
		}
		s.conns[oldestKey].close()
		delete(s.conns, oldestKey)
	}
}

// idle drops connections unused for a while: their PowerShell hosts stop
// and their tokens are forgotten.
func (s *Server) idle(max time.Duration) {
	s.mu.Lock()
	defer s.mu.Unlock()
	for k, c := range s.conns {
		if time.Since(c.used) > max {
			c.close()
			delete(s.conns, k)
		}
	}
}

// Serve listens until ctx ends.
func (s *Server) Serve(ctx context.Context, addr string) error {
	l, err := net.Listen("tcp", addr)
	if err != nil {
		return err
	}
	return s.ServeListener(ctx, l)
}

// ServeListener serves on l (tests pass their own).
func (s *Server) ServeListener(ctx context.Context, l net.Listener) error {
	// No write timeout: a cmdlet may run for minutes.
	srv := &http.Server{Handler: s.Handler(), ReadHeaderTimeout: 10 * time.Second, ReadTimeout: time.Minute,
		IdleTimeout: 2 * time.Minute, TLSConfig: s.TLSConfig()}
	go func() {
		t := time.NewTicker(time.Minute)
		defer t.Stop()
		for {
			select {
			case <-ctx.Done():
				_ = srv.Close()
				s.mu.Lock()
				for k, c := range s.conns {
					c.close()
					delete(s.conns, k)
				}
				s.mu.Unlock()
				return
			case <-t.C:
				s.idle(15 * time.Minute)
			}
		}
	}()
	err := srv.ServeTLS(l, "", "")
	if errors.Is(err, http.ErrServerClosed) {
		return nil
	}
	return err
}

func (s *Server) audit(action, target, detail string, err error) {
	if s.Audit == nil {
		return
	}
	e := auditlog.Entry{Time: time.Now(), Action: action, Target: target, Detail: detail, OK: err == nil, Profile: "worker"}
	if err != nil {
		e.Error = err.Error()
	}
	s.Audit.Write(e)
}

func (s *Server) log(format string, a ...any) {
	if s.Logf != nil {
		s.Logf(format, a...)
	}
}
