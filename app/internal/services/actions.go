package services

import (
	"errors"

	"swissknife-app/internal/actions"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/exoapi"
	"swissknife-app/internal/session"
)

// ActionsService exposes the action catalog (ADR-008): list what can run,
// preview a change, apply the preview.
type ActionsService struct {
	e *engine.Engine
}

// NewEngine builds the session's engine with every built-in action; shared by
// the GUI binding and the CLI.
func NewEngine(s *session.Session) *engine.Engine {
	e := engine.New(s, engine.GraphProvider{}, exoapi.NewProvider())
	e.WrapErr = wrapOpErr
	e.Register(actions.Builtin()...)
	return e
}

func NewActionsService(s *session.Session) *ActionsService {
	return &ActionsService{e: NewEngine(s)}
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
