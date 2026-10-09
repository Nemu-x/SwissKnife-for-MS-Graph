package services

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/packs"
	"swissknife-app/internal/session"
)

// The PowerShell host learns each trusted script with the commands its pack
// declared; the pack list shows them, network/disk ones flagged.
func TestTrustedScriptsCarryTheirDeclaredCommands(t *testing.T) {
	dir := t.TempDir()
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(dir)
	pack := filepath.Join(dir, "actions", "contoso")
	_ = os.MkdirAll(pack, 0o755)
	man := strings.Replace(holdManifest, "cmdlets: [Get-Mailbox, Set-Mailbox]", "cmdlets: [Set-Mailbox, Get-Mailbox, Export-Csv]", 1) +
		"permissions: [\"Exchange: Recipient Management\"]\n"
	_ = os.WriteFile(filepath.Join(pack, packs.ManifestFile), []byte(man), 0o644)
	_ = os.WriteFile(filepath.Join(pack, "hold.ps1"), []byte("param($Mode)"), 0o644)
	x := NewPacksService(sess)
	list, _ := x.List()
	if len(trustedScriptHashes(sess)()) != 0 {
		t.Fatal("an untrusted pack's scripts are not given to the host")
	}
	if _, err := x.Trust("contoso-tools", list[0].Digest); err != nil {
		t.Fatal(err)
	}
	got := trustedScriptHashes(sess)()
	if len(got) != 1 || !reflect.DeepEqual(got[0].Cmdlets, []string{"Export-Csv", "Get-Mailbox", "Set-Mailbox"}) {
		t.Fatalf("trusted scripts %+v", got)
	}
	list, _ = x.List()
	a := list[0].Actions[0]
	if !reflect.DeepEqual(a.Sensitive, []string{"Export-Csv"}) || len(list[0].Permissions) != 1 {
		t.Fatalf("pack info %+v", list[0])
	}
}

func TestSigningKeysNeedAPublisherName(t *testing.T) {
	sess := session.New(auditlog.New(t.TempDir()))
	sess.SetConfigDir(t.TempDir())
	x := NewPacksService(sess)
	const key = "RWQf6LRCGA9i53mlYecO4IzT51TGPpvWucNSCh1CBM0QTaLn73Y7GFO3"
	if _, err := x.AddKey(key, ""); err == nil {
		t.Fatal("a key needs a publisher name")
	}
	if _, err := x.AddKey(key, "SwissKnife Project"); err == nil {
		t.Fatal("a user key cannot take the project's name")
	}
	keys, err := x.AddKey(key, "Contoso IT")
	if err != nil || len(keys) != 1 || keys[0].Name != "Contoso IT" {
		t.Fatalf("keys %+v %v", keys, err)
	}
	if keys, _ = x.AddKey(key, "Contoso Security"); len(keys) != 1 || keys[0].Name != "Contoso Security" {
		t.Fatalf("renaming a key: %+v", keys)
	}
	if keys, _ = x.RemoveKey(key); len(keys) != 0 || len(packs.LoadTrust(trustFile(sess)).Names) != 0 {
		t.Fatalf("removed: %+v", keys)
	}
}
