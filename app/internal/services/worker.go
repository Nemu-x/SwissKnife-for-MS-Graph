package services

import (
	"context"
	"encoding/json"
	"errors"
	"net"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
	"swissknife-app/internal/worker"
)

// The client side of a remote PowerShell worker (ADR-008 D9): when this
// machine cannot run a module family (no PowerShell, a module missing, or
// a Windows-only module on a Mac), the cmdlet calls go to the paired worker.
// The action itself — preview, confirmation, journal, audit — stays here.

// WorkerInfo is the paired worker (worker.json; no secrets in it).
type WorkerInfo struct {
	Addr        string    `json:"addr"`
	Fingerprint string    `json:"fingerprint"`
	Name        string    `json:"name"`
	PairedAt    time.Time `json:"pairedAt"`
	Families    []string  `json:"families"` // from its last health answer
}

var workerMu sync.Mutex

func workerFile(s *session.Session) string { return filepath.Join(s.ConfigDir(), "worker.json") }

func loadWorker(s *session.Session) *WorkerInfo {
	if s.ConfigDir() == "" {
		return nil
	}
	b, err := os.ReadFile(workerFile(s))
	if err != nil {
		return nil
	}
	var w WorkerInfo
	if json.Unmarshal(b, &w) != nil || w.Addr == "" || w.Fingerprint == "" {
		return nil
	}
	return &w
}

func saveWorker(s *session.Session, w *WorkerInfo) error {
	if w == nil {
		err := os.Remove(workerFile(s))
		if errors.Is(err, os.ErrNotExist) {
			return nil
		}
		return err
	}
	b, _ := json.MarshalIndent(w, "", "  ")
	tmp := workerFile(s) + ".tmp"
	if err := os.WriteFile(tmp, b, 0o600); err != nil {
		return err
	}
	return os.Rename(tmp, workerFile(s))
}

func clientIdentity(s *session.Session) (*worker.Identity, error) {
	if s.ConfigDir() == "" {
		return nil, errors.New("config directory is not set")
	}
	return worker.LoadOrCreateIdentity(s.ConfigDir(), "worker-client")
}

// remotes caches one client per worker, so calls reuse their connections.
var remotes = map[string]*worker.Remote{}

// remoteFor returns the paired worker's client and info, or nil.
func remoteFor(s *session.Session) (*worker.Remote, *WorkerInfo) {
	workerMu.Lock()
	defer workerMu.Unlock()
	w := loadWorker(s)
	if w == nil {
		return nil, nil
	}
	key := w.Addr + "|" + w.Fingerprint
	if r := remotes[key]; r != nil {
		return r, w
	}
	id, err := clientIdentity(s)
	if err != nil {
		return nil, nil
	}
	r := worker.NewRemote(w.Addr, w.Fingerprint, id)
	remotes[key] = r
	return r, w
}

func familyModule(family string) string {
	if family == pwsh.FamilyTeams {
		return pwsh.ModuleTeams
	}
	return pwsh.ModuleExchange
}

func hasFamily(w *WorkerInfo, family string) bool {
	if w == nil {
		return false
	}
	for _, f := range w.Families {
		if f == family {
			return true
		}
	}
	return false
}

// psDispatch runs a cmdlet locally when this machine can, else on the
// worker. Pack scripts never leave this machine.
type psDispatch struct {
	s     *session.Session
	local *pwsh.Pool
	det   *pwsh.Detector
	can   func(family string) bool // tests: stands in for the detector
}

func (d *psDispatch) localCan(ctx context.Context, family string) bool {
	if d.can != nil {
		return d.can(family)
	}
	env := d.det.Get(ctx)
	return env.PwshOK() && env.Supports(familyModule(family))
}

func (d *psDispatch) Invoke(env engine.Env, family, cmdlet string, params map[string]any, sel ...string) ([]json.RawMessage, error) {
	if d.localCan(env.Ctx, family) {
		return d.local.Invoke(env, family, cmdlet, params, sel...)
	}
	r, w := remoteFor(d.s)
	if r == nil || !hasFamily(w, family) {
		return d.local.Invoke(env, family, cmdlet, params, sel...) // reports what is missing
	}
	if env.Tokens == nil {
		return nil, session.ErrNotConnected
	}
	req := worker.InvokeRequest{Family: family, Cmdlet: cmdlet, Params: params, Select: sel,
		TenantID: env.TenantID, AppOnly: env.AppOnly, DelegatedOrg: env.DelegatedOrg}
	// Teams signs in with a Graph token too; Exchange only needs the org or
	// the admin's UPN, found here, so the worker never gets a Graph token.
	resources := []string{auth.ResourceGraph, auth.ResourceTeams}
	if family != pwsh.FamilyTeams {
		resources = []string{auth.ResourceExchange}
		if err := exchangeSignIn(env, &req); err != nil {
			return nil, err
		}
	}
	toks := map[string]string{}
	for _, res := range resources {
		t, err := env.Tokens.TokenFor(env.Ctx, res)
		if err != nil {
			return nil, err
		}
		toks[res] = t
	}
	req.Tokens = toks
	out, err := r.Invoke(env.Ctx, req)
	d.s.Record("worker.invoke", cmdlet, "worker="+w.Name+" ("+w.Addr+")", err)
	return out, err
}

// exchangeSignIn fills what Exchange PowerShell signs in with: the
// tenant's initial domain (app-only), the customer org (GDAP) or the
// admin's UPN (delegated). Looked up once per connection.
func exchangeSignIn(env engine.Env, req *worker.InvokeRequest) error {
	if env.Graph == nil {
		return session.ErrNotConnected
	}
	switch {
	case env.AppOnly:
		org, err := graphapi.InitialDomain(env.Ctx, env.Graph)
		if err != nil {
			return err
		}
		req.ExchangeOrg = org
	case env.DelegatedOrg != "":
	default:
		var me struct {
			UPN string `json:"userPrincipalName"`
		}
		if err := env.Graph.Get(env.Ctx, "/me", url.Values{"$select": {"userPrincipalName"}}, &me); err != nil {
			return err
		}
		req.ExchangeUPN = me.UPN
	}
	return nil
}

func (d *psDispatch) RunScript(env engine.Env, family, script string, params map[string]any) ([]json.RawMessage, error) {
	return d.local.RunScript(env, family, script, params)
}

func (d *psDispatch) ResetPacks() { d.local.ResetPacks() }

// workerAware makes a PowerShell backend available when the paired worker
// can run its module family although this machine cannot.
type workerAware struct {
	inner  engine.Provider
	family string
}

func (p workerAware) Backend() engine.Backend { return p.inner.Backend() }

func (p workerAware) Status(s *session.Session) *engine.Reason {
	r := p.inner.Status(s)
	if r == nil || r.Key == "notConnected" {
		return r
	}
	if _, w := remoteFor(s); hasFamily(w, p.family) {
		return nil
	}
	return r
}

// --- binding ---------------------------------------------------------------

// WorkerService pairs this app with a worker.
type WorkerService struct{ s *session.Session }

func NewWorkerService(s *session.Session) *WorkerService { return &WorkerService{s: s} }

// Status returns the paired worker (nil when none).
func (x *WorkerService) Status() *WorkerInfo {
	workerMu.Lock()
	defer workerMu.Unlock()
	return loadWorker(x.s)
}

// Pair connects with the code the worker printed and pins it.
func (x *WorkerService) Pair(addr, code string) (*WorkerInfo, error) {
	addr = strings.TrimSpace(addr)
	if addr == "" || strings.TrimSpace(code) == "" {
		return nil, errors.New("enter the worker's address and its pairing code")
	}
	if _, _, err := net.SplitHostPort(addr); err != nil {
		addr = net.JoinHostPort(strings.Trim(addr, "[]"), "8743")
	}
	id, err := clientIdentity(x.s)
	if err != nil {
		return nil, err
	}
	host, _ := os.Hostname()
	ctx, cancel := context.WithTimeout(x.s.Ctx(), 30*time.Second)
	defer cancel()
	fp, name, err := worker.Pair(ctx, addr, code, "SwissKnife on "+host, id)
	x.s.Record("worker.pair", addr, "", err)
	if err != nil {
		return nil, err
	}
	w := &WorkerInfo{Addr: addr, Fingerprint: fp, Name: name, PairedAt: time.Now()}
	if h, err := worker.NewRemote(addr, fp, id).Health(ctx); err == nil {
		w.Families = h.Families
	}
	workerMu.Lock()
	defer workerMu.Unlock()
	return w, saveWorker(x.s, w)
}

// Check asks the worker what it runs now and remembers it.
func (x *WorkerService) Check() (*WorkerInfo, error) {
	r, w := remoteFor(x.s)
	if r == nil {
		return nil, errors.New("no worker is paired")
	}
	ctx, cancel := context.WithTimeout(x.s.Ctx(), 15*time.Second)
	defer cancel()
	h, err := r.Health(ctx)
	if err != nil {
		return nil, err
	}
	w.Families = h.Families
	w.Name = h.Name
	workerMu.Lock()
	defer workerMu.Unlock()
	return w, saveWorker(x.s, w)
}

// Unpair forgets the worker and asks it to forget this app (best effort:
// an unreachable worker is forgotten here anyway — revoke it there).
func (x *WorkerService) Unpair() error {
	var remoteErr error
	if r, _ := remoteFor(x.s); r != nil {
		ctx, cancel := context.WithTimeout(x.s.Ctx(), 10*time.Second)
		remoteErr = r.Unpair(ctx)
		cancel()
		r.Close()
	}
	workerMu.Lock()
	defer workerMu.Unlock()
	for k := range remotes {
		delete(remotes, k)
	}
	x.s.Record("worker.unpair", "", "", remoteErr)
	return saveWorker(x.s, nil)
}

// ClientFingerprint is this app's certificate fingerprint, to recognise it
// in the worker's client list.
func (x *WorkerService) ClientFingerprint() (string, error) {
	id, err := clientIdentity(x.s)
	if err != nil {
		return "", err
	}
	return id.Fingerprint, nil
}
