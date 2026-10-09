package services

import (
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"swissknife-app/internal/packs"
)

const compromisedWorkflow = `name: security-basics
version: 1.0.0
workflows:
  - id: compromisedUser
    page: security
    label: { en: Compromised user, ru: Взломанная учётка }
    confirmField: user
    fields:
      - { name: user, kind: user, required: true }
      - { name: resetMfa, kind: choice, options: ["yes", "no"], default: "no" }
    steps:
      - action: user.signIn
        with: { user: "{{user}}", state: blocked }
      - action: user.revokeSessions
        with: { user: "{{user}}" }
      - action: user.resetMfa
        with: { user: "{{user}}" }
        when: { input: resetMfa, equals: "yes" }
`

func TestWorkflowPackRunsBuiltInActions(t *testing.T) {
	var calls []string
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		calls = append(calls, r.Method+" "+r.URL.Path)
		switch {
		case r.Method == "GET" && strings.HasPrefix(r.URL.Path, "/users/"):
			_, _ = w.Write([]byte(`{"id":"u1","userPrincipalName":"ann@contoso.com","accountEnabled":true}`))
		default:
			w.WriteHeader(http.StatusNoContent)
		}
	})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	pack := filepath.Join(dir, "actions", "security-basics")
	_ = os.MkdirAll(pack, 0o755)
	_ = os.WriteFile(filepath.Join(pack, packs.ManifestFile), []byte(compromisedWorkflow), 0o644)
	list := packs.Load(packsRoot(sess), packs.Trust{})
	if list[0].Status == packs.Invalid {
		t.Fatalf("pack: %s", list[0].Error)
	}
	_ = packs.SaveTrust(trustFile(sess), packs.Trust{Pinned: map[string]string{"security-basics": list[0].Digest}})

	e := NewEngine(sess)
	const id = "pack.security-basics.compromisedUser"
	m, ok := e.ActionManifest(id)
	if !ok || m.Danger != "destructive" || !m.Workflow {
		t.Fatalf("workflow action %+v", m)
	}
	p, err := e.Plan(id, map[string]string{"user": "ann@contoso.com"})
	if err != nil {
		t.Fatal(err)
	}
	if len(p.Changes) < 2 || p.Changes[0].Step != "user.signIn" || p.ConfirmTarget != "ann@contoso.com" {
		t.Fatalf("plan %+v confirm %q", p.Changes, p.ConfirmTarget)
	}
	for _, ch := range p.Changes {
		if ch.Step == "user.resetMfa" {
			t.Fatal("the MFA step runs only when asked")
		}
	}
	if _, err := e.Apply(p.ID, "ann@contoso.com"); err != nil {
		t.Fatal(err)
	}
	joined := strings.Join(calls, "\n")
	if !strings.Contains(joined, "PATCH /users/u1") || !strings.Contains(joined, "revokeSignInSessions") {
		t.Fatalf("calls:\n%s", joined)
	}
}

func TestWorkflowWithAnUnknownActionMakesThePackInvalid(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	dir := t.TempDir()
	sess.SetConfigDir(dir)
	pack := filepath.Join(dir, "actions", "bad")
	_ = os.MkdirAll(pack, 0o755)
	bad := strings.Replace(compromisedWorkflow, "user.revokeSessions", "user.deleteEverything", 1)
	_ = os.WriteFile(filepath.Join(pack, packs.ManifestFile), []byte(bad), 0o644)
	NewEngine(sess)
	list, _ := NewPacksService(sess).List()
	if len(list) != 1 || list[0].Status != packs.Invalid || !strings.Contains(list[0].Error, "no such action") {
		t.Fatalf("list %+v", list)
	}
}

// The sample workflow packs in the repository match the real catalog.
func TestSampleWorkflowPacksMatchTheCatalog(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {})
	e := NewEngine(sess)
	for _, p := range packs.Load(filepath.Join("..", "..", "..", "packs"), packs.Trust{}) {
		if _, err := workflowActions(e, p, nil); err != nil {
			t.Errorf("%s: %v", p.Manifest.Name, err)
		}
	}
}
