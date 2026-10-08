package services

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/auth"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/exoapi"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// ActionsService exposes the action catalog (ADR-008): list what can run,
// preview a change, apply the preview.
type ActionsService struct {
	e *engine.Engine
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
	e := engine.New(s, engine.GraphProvider{}, exoapi.NewProvider(), pwsh.NewExchangeProvider(psDetector))
	e.PS = pool
	e.Grants = func() map[string]bool { return graphGrants(s) }
	e.WrapErr = wrapOpErr
	e.Register(actions.Builtin()...)
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

func NewActionsService(s *session.Session) *ActionsService {
	return &ActionsService{e: EngineFor(s)}
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
	env := psDetector.Get(context.Background())
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
	err := psDetector.Install(context.Background(), name)
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
