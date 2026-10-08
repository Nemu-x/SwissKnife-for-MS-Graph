package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/packs"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

// Community action packs (ADR-008 D13): loaded from <config>/actions, added
// to the catalog, run as trusted scripts in the PowerShell host.

var (
	packMu sync.Mutex
	loaded []packs.Pack // the packs as last loaded
)

func packsRoot(s *session.Session) string { return filepath.Join(s.ConfigDir(), "actions") }
func trustFile(s *session.Session) string { return filepath.Join(s.ConfigDir(), "packs.json") }

// trustedScriptHashes are the scripts the PowerShell host may run.
func trustedScriptHashes() []string {
	packMu.Lock()
	defer packMu.Unlock()
	var out []string
	for _, p := range loaded {
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
	list := packs.Load(packsRoot(s), packs.LoadTrust(trustFile(s)))
	before := strings.Join(trustedScriptHashes(), ",")
	packMu.Lock()
	loaded = list
	packMu.Unlock()
	e.SetPacks(packActions(list))
	// Hosts know their trusted scripts from the start: restart them only
	// when that set changed (a restart signs every module in again).
	if pool, ok := e.PS.(*pwsh.Pool); ok && strings.Join(trustedScriptHashes(), ",") != before {
		pool.Reset()
	}
	return list
}

func packActions(list []packs.Pack) []engine.Action {
	var out []engine.Action
	for _, p := range list {
		if p.Status == packs.Invalid {
			continue
		}
		status := p.Status
		gate := func() *engine.Reason {
			switch status {
			case packs.Untrusted:
				return &engine.Reason{Key: "packUntrusted"}
			case packs.Changed:
				return &engine.Reason{Key: "packChanged"}
			}
			return nil
		}
		for _, d := range p.Manifest.Actions {
			var fields []engine.Field
			for _, f := range d.Fields {
				fields = append(fields, engine.Field{Name: f.Name, Kind: engine.FieldKind(f.Kind), Required: f.Required,
					Options: f.Options, Default: f.Default, Label: f.Label})
			}
			impl := packImpl{def: d, script: p.Scripts[d.Script], family: pwsh.FamilyExchange, backend: pwsh.BackendExchangePS}
			if d.Module == "teams" {
				impl.family, impl.backend = pwsh.FamilyTeams, pwsh.BackendTeamsPS
			}
			a := engine.Action{
				Manifest: engine.Manifest{ID: "pack." + p.Manifest.Name + "." + d.ID, Page: d.Page, Danger: engine.Danger(d.Danger),
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
	}
	return out
}

// packImpl runs a pack script: param($Mode, $Inputs, $Change), where Mode is
// read, plan or apply. Plan returns objects {target, field, op, before,
// after, ref}; apply gets one of them back as $Change.
type packImpl struct {
	def     packs.ActionDef
	script  string
	family  string
	backend engine.Backend
}

func (p packImpl) Backend() engine.Backend { return p.backend }

func (p packImpl) run(env engine.Env, mode string, in engine.Inputs, change map[string]any) ([]json.RawMessage, error) {
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
			Dir: p.Dir, Status: p.Status, Signer: p.Signer, Error: p.Error, Digest: p.Digest, Actions: []PackActionInfo{}}
		if pi.Name == "" {
			pi.Name = filepath.Base(p.Dir)
		}
		for _, a := range p.Manifest.Actions {
			pi.Actions = append(pi.Actions, PackActionInfo{ID: a.ID, Label: a.Label, Page: a.Page, Danger: a.Danger, Module: a.Module})
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

// Untrust removes a pack's pin.
func (x *PacksService) Untrust(name string) ([]PackInfo, error) {
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
