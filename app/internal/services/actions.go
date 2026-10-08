package services

import (
	"context"
	"errors"

	"swissknife-app/internal/actions"
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
	e.WrapErr = wrapOpErr
	e.Register(actions.Builtin()...)
	return e
}

func NewActionsService(s *session.Session) *ActionsService {
	return &ActionsService{e: NewEngine(s)}
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
