package pwsh

import (
	"context"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/session"
)

// Provider reports whether a PowerShell backend can run here: PowerShell 7
// and the family's module installed in a usable version.
type Provider struct {
	backend engine.Backend
	module  string
	det     *Detector
}

func NewExchangeProvider(det *Detector) *Provider {
	return &Provider{backend: BackendExchangePS, module: ModuleExchange, det: det}
}

func NewTeamsProvider(det *Detector) *Provider {
	return &Provider{backend: BackendTeamsPS, module: ModuleTeams, det: det}
}

func (p *Provider) Backend() engine.Backend { return p.backend }

func (p *Provider) Status(s *session.Session) *engine.Reason {
	if !s.Connected() || s.Tokens() == nil {
		return &engine.Reason{Key: "notConnected"}
	}
	ctx, cancel := context.WithTimeout(s.Ctx(), 40*time.Second)
	defer cancel()
	env := p.det.Get(ctx)
	if !env.PwshOK() {
		return &engine.Reason{Key: "pwshMissing"}
	}
	if !env.Supports(p.module) {
		return &engine.Reason{Key: "moduleMissing", Params: map[string]string{"module": p.module}}
	}
	return nil
}
