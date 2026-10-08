package services

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/engine"
	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/packs"
	"swissknife-app/internal/pwsh"
	"swissknife-app/internal/session"
)

const holdManifest = `name: contoso-tools
version: 1.0.0
actions:
  - id: setHold
    page: users
    danger: write
    module: exo
    script: hold.ps1
    label: { en: Put a mailbox on hold }
    fields:
      - { name: mailbox, kind: user, required: true }
`

type fakeScripts struct {
	calls []map[string]any
	reply []string
}

func (f *fakeScripts) Invoke(engine.Env, string, string, map[string]any, ...string) ([]json.RawMessage, error) {
	return nil, nil
}

func (f *fakeScripts) RunScript(_ engine.Env, family, script string, params map[string]any) ([]json.RawMessage, error) {
	f.calls = append(f.calls, params)
	var out []json.RawMessage
	for _, r := range f.reply {
		out = append(out, json.RawMessage(r))
	}
	return out, nil
}

type readyPS struct{}

func (readyPS) Backend() engine.Backend               { return pwsh.BackendExchangePS }
func (readyPS) Status(*session.Session) *engine.Reason { return nil }

func TestPackActionsNeedTrustAndRunThroughTheScriptHost(t *testing.T) {
	dir := t.TempDir()
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(dir)
	sess.SetClient(graphapi.New(graphapi.StaticToken("t")), "test") // pwsh actions run for a tenant
	pack := filepath.Join(dir, "actions", "contoso")
	_ = os.MkdirAll(pack, 0o755)
	_ = os.WriteFile(filepath.Join(pack, packs.ManifestFile), []byte(holdManifest), 0o644)
	_ = os.WriteFile(filepath.Join(pack, "hold.ps1"), []byte("param($Mode, $Inputs, $Change)"), 0o644)

	e := engine.New(sess, readyPS{})
	fake := &fakeScripts{}
	e.PS = fake
	list := packs.Load(packsRoot(sess), packs.LoadTrust(trustFile(sess)))
	e.SetPacks(packActions(sess, list))

	entry := func() engine.CatalogEntry {
		for _, c := range e.Catalog() {
			if c.ID == "pack.contoso-tools.setHold" {
				return c
			}
		}
		t.Fatal("pack action missing from the catalog")
		return engine.CatalogEntry{}
	}
	if c := entry(); c.Available || c.Reason == nil || c.Reason.Key != "packUntrusted" || c.Label["en"] == "" || c.Pack != "contoso-tools" {
		t.Fatalf("untrusted pack: %+v", c)
	}

	// The operator trusts exactly what was reviewed.
	if _, err := NewPacksService(sess).Trust("contoso-tools", "0000"); err == nil {
		t.Fatal("a digest that is not the folder's must be refused")
	}
	if err := packs.SaveTrust(trustFile(sess), packs.Trust{Pinned: map[string]string{"contoso-tools": list[0].Digest}}); err != nil {
		t.Fatal(err)
	}
	e.SetPacks(packActions(sess, packs.Load(packsRoot(sess), packs.LoadTrust(trustFile(sess)))))
	if c := entry(); !c.Available {
		t.Fatalf("trusted pack: %+v", c)
	}

	fake.reply = []string{`{"target":"ann@contoso.com","field":"hold","op":"set","before":"off","after":"on","ref":{"id":"m1","confirmRef":"x"}}`}
	p, err := e.Plan("pack.contoso-tools.setHold", engine.Inputs{"mailbox": "ann@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if p.Changes[0].Ref[engine.ConfirmRef] != "" || p.ConfirmTarget != "" {
		t.Fatal("a pack must not choose its own confirmation")
	}
	if _, err := e.Apply(p.ID, ""); err != nil {
		t.Fatal(err)
	}
	last := fake.calls[len(fake.calls)-1]
	ch, _ := last["Change"].(map[string]any)
	if last["Mode"] != "apply" || ch["target"] != "ann@contoso.com" {
		t.Fatalf("apply call %+v", last)
	}

	// Editing the script after trusting it makes the action unavailable.
	_ = os.WriteFile(filepath.Join(pack, "hold.ps1"), []byte("Remove-Mailbox *"), 0o644)
	e.SetPacks(packActions(sess, packs.Load(packsRoot(sess), packs.LoadTrust(trustFile(sess)))))
	if c := entry(); c.Available || c.Reason.Key != "packChanged" {
		t.Fatalf("changed pack: %+v", c)
	}
	if _, err := e.Plan("pack.contoso-tools.setHold", engine.Inputs{"mailbox": "x"}); err == nil || !strings.Contains(err.Error(), "unavailable") {
		t.Fatalf("a changed pack must not run: %v", err)
	}
}

const readManifest = `name: contoso-read
version: 1.0.0
actions:
  - id: holds
    page: reports
    danger: read
    module: exo
    script: holds.ps1
    label: { en: Holds }
`

// A pack's "read" runs a script: read-only mode, a read-only profile and a
// group scope all stop it, and every run is audited.
func TestPackReadsAreGuardedLikeWrites(t *testing.T) {
	dir := t.TempDir()
	logDir := t.TempDir()
	sess := session.New(auditlog.New(logDir))
	sess.SetConfigDir(dir)
	sess.SetClient(graphapi.New(graphapi.StaticToken("t")), "test")
	pack := filepath.Join(dir, "actions", "r")
	_ = os.MkdirAll(pack, 0o755)
	_ = os.WriteFile(filepath.Join(pack, packs.ManifestFile), []byte(readManifest), 0o644)
	_ = os.WriteFile(filepath.Join(pack, "holds.ps1"), []byte("param($Mode)"), 0o644)
	list := packs.Load(packsRoot(sess), packs.Trust{})
	_ = packs.SaveTrust(trustFile(sess), packs.Trust{Pinned: map[string]string{"contoso-read": list[0].Digest}})

	e := engine.New(sess, readyPS{})
	fake := &fakeScripts{reply: []string{`{"name":"Ann"}`}}
	e.PS = fake
	e.SetPacks(packActions(sess, packs.Load(packsRoot(sess), packs.LoadTrust(trustFile(sess)))))
	const id = "pack.contoso-read.holds"

	if res, err := e.Run(t.Context(), id, nil); err != nil || len(res.Rows) != 1 {
		t.Fatalf("trusted read: %+v %v", res, err)
	}
	sess.SetReadOnly(true)
	if _, err := e.Run(t.Context(), id, nil); err == nil {
		t.Fatal("read-only mode must stop pack scripts")
	}
	sess.SetReadOnly(false)
	sess.SetPolicy("p", session.Policy{MaxDanger: "read"})
	if _, err := e.Run(t.Context(), id, nil); err == nil {
		t.Fatal("a read-only profile must not run pack scripts")
	}
	sess.SetPolicy("p", session.Policy{AllowedGroups: []string{"g1"}})
	if _, err := e.Run(t.Context(), id, nil); err == nil || !strings.Contains(err.Error(), "action-pack") {
		t.Fatalf("a group scope must refuse packs: %v", err)
	}
	if len(fake.calls) != 1 {
		t.Fatalf("refused runs must not reach PowerShell: %d", len(fake.calls))
	}
	b, _ := os.ReadFile(sess.Audit.Path())
	if !strings.Contains(string(b), "pack.run") {
		t.Fatal("pack runs must be audited")
	}
}
