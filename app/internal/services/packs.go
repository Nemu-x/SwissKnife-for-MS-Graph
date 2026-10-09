package services

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/packs"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// Community action packs (ADR-008 D13): loaded from <config>/actions, added
// to the catalog, run as trusted scripts in the PowerShell host.

var (
	packMu  sync.Mutex
	loaded  = map[*session.Session][]packs.Pack{} // the packs as last loaded, per session
	trustMu sync.Mutex                            // packs.json read-modify-write
)

func packsRoot(s *session.Session) string { return filepath.Join(s.ConfigDir(), "actions") }
func trustFile(s *session.Session) string { return filepath.Join(s.ConfigDir(), "packs.json") }

// trustedScriptHashes are the scripts the session's pack hosts may run.
func trustedScriptHashes(s *session.Session) func() []string {
	return func() []string {
		packMu.Lock()
		defer packMu.Unlock()
		return scriptHashesLocked(s)
	}
}

func scriptHashesLocked(s *session.Session) []string {
	var out []string
	for _, p := range loaded[s] {
		if !p.Usable() {
			continue
		}
		for _, text := range p.Scripts {
			out = append(out, pwsh.ScriptHash(text))
		}
	}
	sort.Strings(out)
	return out
}

// reloadPacks reads the packs again, swaps their actions in the engine and
// restarts PowerShell hosts so the trusted scripts take effect.
func reloadPacks(s *session.Session, e *engine.Engine) []packs.Pack {
	if s.ConfigDir() == "" {
		return nil
	}
	// One reload at a time: the engine's actions, the trusted scripts and
	// the hosts must all come from the same load.
	packMu.Lock()
	defer packMu.Unlock()
	list := packs.Load(packsRoot(s), packs.LoadTrust(trustFile(s)))
	before := strings.Join(scriptHashesLocked(s), ",")
	loaded[s] = list
	e.SetPacks(packActions(s, e, list))
	// Pack hosts know their trusted scripts from the start: replace them
	// (at their next use) only when that set changed.
	if pool, ok := e.PS.(interface{ ResetPacks() }); ok && strings.Join(scriptHashesLocked(s), ",") != before {
		pool.ResetPacks()
	}
	return list
}

func packActions(s *session.Session, e *engine.Engine, list []packs.Pack) []engine.Action {
	var out []engine.Action
	for i := range list {
		// Workflows are checked against the catalog: an unknown action or
		// input makes the whole pack invalid, with the reason shown.
		if list[i].Status != packs.Invalid {
			if _, err := workflowActions(e, list[i], nil); err != nil {
				list[i].Status, list[i].Error = packs.Invalid, err.Error()
			}
		}
	}
	for _, p := range list {
		if p.Status == packs.Invalid {
			continue
		}
		status := p.Status
		check := folderCheck(p.Dir, p.Digest)
		gate := func() *engine.Reason {
			// A trusted folder edited since the load counts as changed, even
			// before the next reload (what runs is the loaded, trusted text).
			if (status == packs.Signed || status == packs.Trusted) && !check() {
				return &engine.Reason{Key: "packChanged"}
			}
			switch status {
			case packs.Untrusted:
				return &engine.Reason{Key: "packUntrusted"}
			case packs.Changed:
				return &engine.Reason{Key: "packChanged"}
			case packs.Disabled:
				return &engine.Reason{Key: "packDisabled"}
			}
			return nil
		}
		for _, d := range p.Manifest.Actions {
			var fields []engine.Field
			for _, f := range d.Fields {
				fields = append(fields, engine.Field{Name: f.Name, Kind: engine.FieldKind(f.Kind), Required: f.Required,
					Options: f.Options, Default: f.Default, Label: f.Label})
			}
			id := "pack." + p.Manifest.Name + "." + d.ID
			impl := packImpl{s: s, id: id, def: d, script: p.Scripts[d.Script], family: pwsh.FamilyExchange, backend: pwsh.BackendExchangePS}
			if d.Module == "teams" {
				impl.family, impl.backend = pwsh.FamilyTeams, pwsh.BackendTeamsPS
			}
			a := engine.Action{
				Manifest: engine.Manifest{ID: id, Page: d.Page, Danger: engine.Danger(d.Danger),
					Fields: fields, ConfirmField: d.ConfirmField, Label: d.Label, Hint: d.Hint, Pack: p.Manifest.Name},
				Gate: gate,
			}
			if d.Danger == "read" {
				a.Impls = []engine.Impl{engine.ReadImpl(packReader{impl})}
			} else {
				a.Impls = []engine.Impl{impl}
			}
			out = append(out, a)
		}
		wf, _ := workflowActions(e, p, gate)
		out = append(out, wf...)
	}
	return out
}

// workflowActions builds a pack's workflows as catalog actions.
func workflowActions(e *engine.Engine, p packs.Pack, gate func() *engine.Reason) ([]engine.Action, error) {
	var out []engine.Action
	for _, w := range p.Manifest.Workflows {
		var fields []engine.Field
		for _, f := range w.Fields {
			fields = append(fields, engine.Field{Name: f.Name, Kind: engine.FieldKind(f.Kind), Required: f.Required,
				Options: f.Options, Default: f.Default, Label: f.Label})
		}
		var steps []engine.WorkflowStep
		for _, st := range w.Steps {
			step := engine.WorkflowStep{Action: st.Action, With: st.With, OnError: st.OnError}
			if len(st.When) > 0 {
				step.When = &engine.StepCondition{Input: st.When["input"], Equals: st.When["equals"]}
			}
			steps = append(steps, step)
		}
		danger, perms, err := e.ValidateWorkflow(fields, w.ConfirmField, steps)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", w.ID, err)
		}
		trust := gate
		out = append(out, engine.Action{
			Manifest: engine.Manifest{ID: "pack." + p.Manifest.Name + "." + w.ID, Page: w.Page, Danger: danger,
				Fields: fields, ConfirmField: w.ConfirmField, Label: w.Label, Hint: w.Hint, Pack: p.Manifest.Name, Workflow: true,
				Permissions: perms},
			Impls: []engine.Impl{e.NewWorkflow(steps)},
			// Trusted, and every step can run now.
			Gate: func() *engine.Reason {
				if trust != nil {
					if r := trust(); r != nil {
						return r
					}
				}
				return e.WorkflowGate(steps)
			},
		})
	}
	return out, nil
}

// packImpl runs a pack script: param($Mode, $Inputs, $Change), where Mode is
// read, plan or apply. Plan returns objects {target, field, op, before,
// after, ref}; apply gets one of them back as $Change.
type packImpl struct {
	s       *session.Session
	id      string
	def     packs.ActionDef
	script  string
	family  string
	backend engine.Backend
}

func (p packImpl) Backend() engine.Backend { return p.backend }

func (p packImpl) run(env engine.Env, mode string, in engine.Inputs, change map[string]any) (out []json.RawMessage, err error) {
	// Every mode runs the pack's code — a preview or a "read" included — so
	// read-only mode stops all of them, and each run is audited.
	if env.ReadOnly {
		return nil, session.ErrReadOnly
	}
	if p.s != nil {
		defer func() {
			p.s.Record("pack.run", p.id, "mode="+mode+" script="+pwsh.ScriptHash(p.script)[:16], err)
		}()
	}
	runner, ok := env.PS.(engine.ScriptRunner)
	if !ok || env.PS == nil {
		return nil, errors.New("PowerShell is not available")
	}
	inputs := map[string]any{}
	for k, v := range in {
		inputs[k] = v
	}
	params := map[string]any{"Mode": mode, "Inputs": inputs}
	if change != nil {
		params["Change"] = change
	}
	return runner.RunScript(env, p.family, p.script, params)
}

const maxPackChanges = 500

func (p packImpl) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	out, err := p.run(env, "plan", in, nil)
	if err != nil {
		return nil, err
	}
	if len(out) > maxPackChanges {
		return nil, fmt.Errorf("the pack planned %d changes; at most %d at a time", len(out), maxPackChanges)
	}
	var changes []engine.Change
	for _, raw := range out {
		var o map[string]any
		if json.Unmarshal(raw, &o) != nil {
			return nil, errors.New("the pack's plan must return objects with target, op, before, after")
		}
		ch := engine.Change{Target: str(o["target"]), Field: str(o["field"]), Op: str(o["op"]), Before: str(o["before"]), After: str(o["after"])}
		switch ch.Op {
		case "set", "add", "remove", "none":
		default:
			return nil, fmt.Errorf("the pack returned op %q (set, add, remove or none)", ch.Op)
		}
		if ch.Target == "" {
			return nil, errors.New("the pack returned a change without a target")
		}
		if ref, ok := o["ref"].(map[string]any); ok {
			ch.Ref = map[string]string{}
			for k, v := range ref {
				if k != engine.ConfirmRef { // a pack cannot choose its own confirmation
					ch.Ref[k] = str(v)
				}
			}
		}
		changes = append(changes, ch)
	}
	return changes, nil
}

func (p packImpl) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	ref := map[string]any{}
	for k, v := range ch.Ref {
		ref[k] = v
	}
	_, err := p.run(env, "apply", in, map[string]any{
		"target": ch.Target, "field": ch.Field, "op": ch.Op, "before": ch.Before, "after": ch.After, "ref": ref,
	})
	return err
}

type packReader struct{ packImpl }

func (r packReader) Read(env engine.Env, in engine.Inputs) (*engine.ReadResult, error) {
	out, err := r.run(env, "read", in, nil)
	if err != nil {
		return nil, err
	}
	res := &engine.ReadResult{Columns: r.def.Columns}
	for _, raw := range out {
		row := engine.Row{}
		var o map[string]any
		if json.Unmarshal(raw, &o) == nil {
			for k, v := range o {
				row[k] = str(v)
			}
		} else {
			var v any
			_ = json.Unmarshal(raw, &v)
			row["value"] = str(v)
		}
		res.Rows = append(res.Rows, row)
	}
	if len(res.Columns) == 0 && len(res.Rows) > 0 {
		for k := range res.Rows[0] {
			res.Columns = append(res.Columns, k)
		}
		sort.Strings(res.Columns)
	}
	return res, nil
}

// str renders a value from PowerShell for a table cell or a plan line.
func str(v any) string {
	switch t := v.(type) {
	case nil:
		return ""
	case string:
		return t
	case float64:
		return strconv.FormatFloat(t, 'f', -1, 64)
	case bool:
		return strconv.FormatBool(t)
	default:
		b, _ := json.Marshal(t)
		return string(b)
	}
}

// --- binding ---------------------------------------------------------------

// PackInfo is a pack as the Settings page shows it.
type PackInfo struct {
	Name        string           `json:"name"`
	Version     string           `json:"version"`
	Author      string           `json:"author"`
	Description string           `json:"description"`
	Dir         string           `json:"dir"`
	Status      packs.Status     `json:"status"`
	Signer      string           `json:"signer,omitempty"`
	Error       string           `json:"error,omitempty"`
	Digest      string           `json:"digest"`
	Actions     []PackActionInfo `json:"actions"`
	Workflows   []PackFlowInfo   `json:"workflows"`
}

// PackFlowInfo lists one workflow of a pack and the actions it chains.
type PackFlowInfo struct {
	ID    string            `json:"id"`
	Label map[string]string `json:"label"`
	Page  string            `json:"page"`
	Steps []string          `json:"steps"`
}

// PackActionInfo lists one action of a pack.
type PackActionInfo struct {
	ID     string            `json:"id"`
	Label  map[string]string `json:"label"`
	Page   string            `json:"page"`
	Danger string            `json:"danger"`
	Module string            `json:"module"`
}

// PacksService manages packs and their trust.
type PacksService struct{ s *session.Session }

func NewPacksService(s *session.Session) *PacksService { return &PacksService{s: s} }

func (x *PacksService) infos(list []packs.Pack) []PackInfo {
	out := []PackInfo{}
	for _, p := range list {
		pi := PackInfo{Name: p.Manifest.Name, Version: p.Manifest.Version, Author: p.Manifest.Author, Description: p.Manifest.Description,
			Dir: p.Dir, Status: p.Status, Signer: p.Signer, Error: p.Error, Digest: p.Digest, Actions: []PackActionInfo{}, Workflows: []PackFlowInfo{}}
		if pi.Name == "" {
			pi.Name = filepath.Base(p.Dir)
		}
		for _, a := range p.Manifest.Actions {
			pi.Actions = append(pi.Actions, PackActionInfo{ID: a.ID, Label: a.Label, Page: a.Page, Danger: a.Danger, Module: a.Module})
		}
		for _, w := range p.Manifest.Workflows {
			f := PackFlowInfo{ID: w.ID, Label: w.Label, Page: w.Page, Steps: []string{}}
			for _, st := range w.Steps {
				f.Steps = append(f.Steps, st.Action)
			}
			pi.Workflows = append(pi.Workflows, f)
		}
		out = append(out, pi)
	}
	return out
}

// List reloads the packs folder and reports every pack.
func (x *PacksService) List() ([]PackInfo, error) {
	if x.s.ConfigDir() == "" {
		return nil, errors.New("config directory is not set")
	}
	return x.infos(reloadPacks(x.s, EngineFor(x.s))), nil
}

// Trust pins a pack's exact contents. digest is what the operator reviewed:
// if the folder changed since, nothing is trusted.
func (x *PacksService) Trust(name, digest string) ([]PackInfo, error) {
	trustMu.Lock()
	defer trustMu.Unlock()
	list := packs.Load(packsRoot(x.s), packs.LoadTrust(trustFile(x.s)))
	for _, p := range list {
		if p.Manifest.Name != name || p.Status == packs.Invalid {
			continue
		}
		if p.Digest != digest {
			return nil, errors.New("the pack changed since you reviewed it — review it again")
		}
		t := packs.LoadTrust(trustFile(x.s))
		t.Pinned[name] = digest
		if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
			return nil, err
		}
		x.s.Record("pack.trust", name, "digest="+digest, nil)
		return x.List()
	}
	return nil, fmt.Errorf("no loadable pack named %q", name)
}

// Disable turns a pack off, even a signed one; Enable undoes it.
func (x *PacksService) Disable(name string) ([]PackInfo, error) { return x.setDisabled(name, true) }

func (x *PacksService) Enable(name string) ([]PackInfo, error) { return x.setDisabled(name, false) }

func (x *PacksService) setDisabled(name string, off bool) ([]PackInfo, error) {
	trustMu.Lock()
	defer trustMu.Unlock()
	t := packs.LoadTrust(trustFile(x.s))
	kept := []string{}
	for _, n := range t.Disabled {
		if n != name {
			kept = append(kept, n)
		}
	}
	if off {
		kept = append(kept, name)
	}
	t.Disabled = kept
	if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
		return nil, err
	}
	x.s.Record("pack.disable", name, fmt.Sprintf("disabled=%v", off), nil)
	return x.List()
}

// Untrust removes a pack's pin.
func (x *PacksService) Untrust(name string) ([]PackInfo, error) {
	trustMu.Lock()
	defer trustMu.Unlock()
	t := packs.LoadTrust(trustFile(x.s))
	delete(t.Pinned, name)
	if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
		return nil, err
	}
	x.s.Record("pack.untrust", name, "", nil)
	return x.List()
}

// Keys lists the extra trusted signing keys.
func (x *PacksService) Keys() []string {
	k := packs.LoadTrust(trustFile(x.s)).Keys
	if k == nil {
		k = []string{}
	}
	return k
}

// AddKey trusts an author's minisign public key (the base64 line).
func (x *PacksService) AddKey(key string) ([]string, error) {
	trustMu.Lock()
	defer trustMu.Unlock()
	key = strings.TrimSpace(key)
	if i := strings.LastIndex(key, "\n"); i >= 0 { // a whole .pub file was pasted
		key = strings.TrimSpace(key[i+1:])
	}
	if err := packs.CheckKey(key); err != nil {
		return nil, fmt.Errorf("not a minisign public key: %w", err)
	}
	t := packs.LoadTrust(trustFile(x.s))
	for _, k := range t.Keys {
		if k == key {
			return t.Keys, nil
		}
	}
	t.Keys = append(t.Keys, key)
	if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
		return nil, err
	}
	x.s.Record("pack.addKey", key, "", nil)
	reloadPacks(x.s, EngineFor(x.s))
	return t.Keys, nil
}

// RemoveKey stops trusting a signing key.
func (x *PacksService) RemoveKey(key string) ([]string, error) {
	trustMu.Lock()
	defer trustMu.Unlock()
	t := packs.LoadTrust(trustFile(x.s))
	kept := []string{}
	for _, k := range t.Keys {
		if k != key {
			kept = append(kept, k)
		}
	}
	t.Keys = kept
	if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
		return nil, err
	}
	x.s.Record("pack.removeKey", key, "", nil)
	reloadPacks(x.s, EngineFor(x.s))
	return kept, nil
}

// OpenFolder shows the packs folder (created if missing).
func (x *PacksService) OpenFolder() error {
	if x.s.ConfigDir() == "" {
		return errors.New("config directory is not set")
	}
	dir := packsRoot(x.s)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return err
	}
	return revealInFolder(dir)
}

// folderCheck reports whether a pack folder still has the digest it was
// loaded with; the answer is reused for a moment (the catalog asks per action).
func folderCheck(dir, digest string) func() bool {
	var mu sync.Mutex
	var at time.Time
	var same bool
	return func() bool {
		mu.Lock()
		defer mu.Unlock()
		if time.Since(at) > 2*time.Second {
			d, err := packs.Digest(dir)
			same, at = err == nil && d == digest, time.Now()
		}
		return same
	}
}

// --- Action Hub --------------------------------------------------------------

// HubPack is a catalog entry with what is installed here.
type HubPack struct {
	packs.HubEntry
	Installed string `json:"installed,omitempty"` // installed version
	Update    bool   `json:"update"`              // the hub has another version
}

// HubView is the catalog as the Settings page shows it.
type HubView struct {
	URL   string    `json:"url"`
	Packs []HubPack `json:"packs"`
}

func (x *PacksService) hubURL() string {
	if h := packs.LoadTrust(trustFile(x.s)).Hub; h != "" {
		return h
	}
	return packs.DefaultHub
}

// HubCatalog reads the Action Hub's index.
func (x *PacksService) HubCatalog() (*HubView, error) {
	ctx, cancel := context.WithTimeout(x.s.Ctx(), time.Minute)
	defer cancel()
	hub := x.hubURL()
	idx, err := packs.FetchIndex(ctx, hub)
	if err != nil {
		return nil, err
	}
	installed := map[string]string{}
	for _, p := range packs.Load(packsRoot(x.s), packs.LoadTrust(trustFile(x.s))) {
		installed[p.Manifest.Name] = p.Manifest.Version
	}
	view := &HubView{URL: hub, Packs: []HubPack{}}
	for _, e := range idx.Packs {
		v, ok := installed[e.Name]
		view.Packs = append(view.Packs, HubPack{HubEntry: e, Installed: v, Update: ok && v != e.Version})
	}
	return view, nil
}

// HubInstall downloads a pack from the hub (or updates it). It still needs
// a trusted signature or the operator's review before it runs.
func (x *PacksService) HubInstall(name string) ([]PackInfo, error) {
	ctx, cancel := context.WithTimeout(x.s.Ctx(), 2*time.Minute)
	defer cancel()
	hub := x.hubURL()
	idx, err := packs.FetchIndex(ctx, hub)
	if err != nil {
		return nil, err
	}
	for _, e := range idx.Packs {
		if e.Name != name {
			continue
		}
		dir, err := packs.Install(ctx, hub, packsRoot(x.s), e)
		x.s.Record("pack.install", name, "hub="+hub+" version="+e.Version+" digest="+e.Digest, err)
		if err != nil {
			return nil, err
		}
		_ = dir
		return x.List()
	}
	return nil, fmt.Errorf("the hub has no pack named %q", name)
}

// SetHub changes the Action Hub address ("" restores the default).
func (x *PacksService) SetHub(hub string) (string, error) {
	hub = strings.TrimSpace(hub)
	if hub != "" && !strings.HasPrefix(hub, "https://") {
		return "", errors.New("the hub address must start with https://")
	}
	trustMu.Lock()
	defer trustMu.Unlock()
	t := packs.LoadTrust(trustFile(x.s))
	t.Hub = hub
	if err := packs.SaveTrust(trustFile(x.s), t); err != nil {
		return "", err
	}
	x.s.Record("pack.hub", hub, "", nil)
	return x.hubURL(), nil
}
