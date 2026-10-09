package services

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/ldapx"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/exoapi"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// ActionsService exposes the action catalog (ADR-008): list what can run,
// preview a change, apply the preview.
type ActionsService struct {
	e     *engine.Engine
	s     *session.Session
	store *secrets.Store // saved profiles, for cross-tenant reads
}

// psDetector is shared by every engine in the process: detecting PowerShell
// spawns pwsh, and the installed modules do not differ per session.
var psDetector = pwsh.NewDetector()

// NewEngine builds the session's engine with every built-in action; shared by
// the GUI binding and the CLI. PowerShell hosts started for the session stop
// when it disconnects.
func NewEngine(s *session.Session) *engine.Engine {
	pool := pwsh.NewPool(psDetector, actions.PowerShellCmdlets())
	// Closing waits for a sign-in that may be in progress; Disconnect is a UI
	// call and must not hang on it.
	s.OnDisconnect(func(prev *graphapi.Client) { go pool.CloseFor(prev) })
	dir := func() *ldapx.Client { return directoryFor(s) }
	e := engine.New(s, engine.GraphProvider{}, exoapi.NewProvider(),
		workerAware{pwsh.NewExchangeProvider(psDetector), pwsh.FamilyExchange},
		workerAware{pwsh.NewTeamsProvider(psDetector), pwsh.FamilyTeams},
		engine.LDAPProvider{Get: dir}, engine.LDAPProvider{Get: dir, Secure: true}, engine.WorkflowProvider{})
	e.LDAP = dir
	// Cmdlets run here when this machine can, else on a paired worker.
	e.PS = &psDispatch{s: s, local: pool, det: psDetector}
	pool.Scripts = trustedScriptHashes(s)
	e.Grants = cachedGrants(s)
	e.WrapErr = wrapOpErr
	// The Exchange sign-in name cached for the worker ends with the connection.
	s.OnDisconnect(func(prev *graphapi.Client) { signIns.Delete(prev) })
	// A cross-tenant read belongs to the connection that started it.
	s.OnDisconnect(func(*graphapi.Client) { s.Ops.CancelKind(ops.KindFanOut) })
	// Service writes outside the catalog check their targets through the
	// session guards; the engine does the lookup.
	s.SetScopeCheck(func(t session.Target) error {
		return engineErr(e.CheckTarget(e.Env(s.Ctx()), engine.FieldKind(t.Kind), t.ID))
	})
	e.Register(actions.Builtin()...)
	e.Register(snapshotRestoreAction(s))
	reloadPacks(s, e)
	return e
}

// engines holds one engine per session, shared by the Actions binding and
// the playbooks, so PowerShell hosts and probes are not duplicated.
var engines sync.Map // *session.Session → *engine.Engine

// EngineFor returns the session's shared engine.
func EngineFor(s *session.Session) *engine.Engine {
	if e, ok := engines.Load(s); ok {
		return e.(*engine.Engine)
	}
	e, _ := engines.LoadOrStore(s, NewEngine(s))
	return e.(*engine.Engine)
}

func NewActionsService(s *session.Session, store *secrets.Store) *ActionsService {
	return &ActionsService{e: EngineFor(s), s: s, store: store}
}

// cachedGrants reads the token's grants once per connection (and again after
// five minutes, when a consent may have changed), so listing the catalog does
// not wait on the token broker every time.
func cachedGrants(s *session.Session) func() map[string]bool {
	var (
		mu   sync.Mutex
		conn *graphapi.Client
		at   time.Time
		have map[string]bool
	)
	return func() map[string]bool {
		c, err := s.Client()
		if err != nil {
			return nil
		}
		mu.Lock()
		defer mu.Unlock()
		// A failed read (nil) is not cached: the next listing tries again.
		if c != conn || have == nil || time.Since(at) > 5*time.Minute {
			conn, at, have = c, time.Now(), graphGrants(s)
		}
		return have
	}
}

// graphGrants reads the permissions from the connection's Graph token:
// "roles" for app-only, "scp" for delegated. The token is already cached by
// the broker; its signature is not checked here — it is our own token, and
// only its claims are read for a hint.
func graphGrants(s *session.Session) map[string]bool {
	b := s.Tokens()
	if b == nil {
		return nil
	}
	ctx, cancel := context.WithTimeout(s.Ctx(), 10*time.Second)
	defer cancel()
	tok, err := b.TokenFor(ctx, auth.ResourceGraph)
	if err != nil {
		return nil
	}
	parts := strings.Split(tok, ".")
	if len(parts) != 3 {
		return nil
	}
	raw, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return nil
	}
	var claims struct {
		Roles []string `json:"roles"`
		Scp   string   `json:"scp"`
	}
	if json.Unmarshal(raw, &claims) != nil {
		return nil
	}
	have := map[string]bool{}
	for _, r := range claims.Roles {
		have[r] = true
	}
	for _, sc := range strings.Fields(claims.Scp) {
		have[sc] = true
	}
	return have
}

// PowerShellStatus reports PowerShell 7 and the modules the app can use.
func (a *ActionsService) PowerShellStatus() PowerShellStatus {
	env := psDetector.Get(a.s.Ctx())
	return PowerShellStatus{
		Installed: env.PwshOK(), Exe: env.Exe, Version: env.Version,
		Modules: []ModuleStatus{
			{Name: pwsh.ModuleExchange, Version: env.Modules[pwsh.ModuleExchange], Usable: env.Supports(pwsh.ModuleExchange)},
			{Name: pwsh.ModuleTeams, Version: env.Modules[pwsh.ModuleTeams], Usable: env.Supports(pwsh.ModuleTeams)},
		},
	}
}

// InstallModule installs a supported module for the current user (minutes).
func (a *ActionsService) InstallModule(name string) (PowerShellStatus, error) {
	err := psDetector.Install(a.s.Ctx(), name)
	return a.PowerShellStatus(), err
}

// installingPwsh lets one PowerShell installation run at a time.
var installingPwsh atomic.Bool

// PowerShellInstallMethod says whether the app can install PowerShell 7 here
// and, if not, what to run by hand.
func (a *ActionsService) PowerShellInstallMethod() pwsh.InstallMethod { return pwsh.Method() }

// InstallPowerShell downloads and installs PowerShell 7 (checked against the
// published checksum; the system asks for admin rights once). Progress
// arrives as "pwsh:install" events.
func (a *ActionsService) InstallPowerShell() (PowerShellStatus, error) {
	if !installingPwsh.CompareAndSwap(false, true) {
		return a.PowerShellStatus(), errors.New("PowerShell is already being installed")
	}
	defer installingPwsh.Store(false)
	err := pwsh.InstallPowerShell(a.s.Ctx(), func(stage string, pct int) {
		emitEvent(a.s.Ctx(), "pwsh:install", map[string]any{"stage": stage, "pct": pct})
	})
	a.s.Record("powershell.install", pwsh.Method().How, "", err)
	psDetector.Invalidate()
	return a.PowerShellStatus(), err
}

// RefreshPowerShell re-detects after the operator installed something by hand.
func (a *ActionsService) RefreshPowerShell() PowerShellStatus {
	psDetector.Invalidate()
	return a.PowerShellStatus()
}

// PowerShellStatus is the Settings view of the PowerShell backends.
type PowerShellStatus struct {
	Installed bool           `json:"installed"`
	Exe       string         `json:"exe"`
	Version   string         `json:"version"`
	Modules   []ModuleStatus `json:"modules"`
}

type ModuleStatus struct {
	Name    string `json:"name"`
	Version string `json:"version"`
	Usable  bool   `json:"usable"`
}

// Catalog lists every action with its availability for the current session.
func (a *ActionsService) Catalog() []engine.CatalogEntry { return a.e.Catalog() }

// Plan previews an action without writing.
func (a *ActionsService) Plan(actionID string, inputs map[string]string) (*engine.Plan, error) {
	p, err := a.e.Plan(actionID, inputs)
	return p, engineErr(err)
}

// Apply runs a previewed plan; confirm is the retyped target for destructive
// actions and ignored otherwise.
func (a *ActionsService) Apply(planID, confirm string) (*engine.Result, error) {
	r, err := a.e.Apply(planID, confirm)
	return r, engineErr(err)
}

// Run executes a read action and returns its rows.
func (a *ActionsService) Run(actionID string, inputs map[string]string) (*engine.ReadResult, error) {
	// Bound by the app's lifetime and a timeout: a read must not hang the UI.
	ctx, cancel := context.WithTimeout(a.s.Ctx(), 5*time.Minute)
	defer cancel()
	r, err := a.e.Run(ctx, actionID, inputs)
	return r, engineErr(err)
}

// Cancel aborts a running apply.
func (a *ActionsService) Cancel(opID string) { a.e.Cancel(opID) }

// engineErr puts engine refusals into the operr envelope so the UI can
// translate them by code.
func engineErr(err error) error {
	var ee *engine.Error
	if errors.As(err, &ee) {
		return &OpError{Code: ee.Code, Message: ee.Msg}
	}
	return err
}

// targetInScope applies the connected profile's group scope to a target
// named outside the catalog (playbooks).
func targetInScope(s *session.Session, kind engine.FieldKind, v string) error {
	e := EngineFor(s)
	return engineErr(e.CheckTarget(e.Env(s.Ctx()), kind, v))
}
