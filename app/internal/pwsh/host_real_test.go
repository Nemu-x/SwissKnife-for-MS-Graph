package pwsh

import (
	"context"
	"encoding/json"
	"errors"
	"strings"
	"testing"
	"time"
)

// TestRealHostScript runs the embedded host.ps1 under a real PowerShell 7
// when one is installed (CI runners have it): the protocol, splatting of
// injection-looking values, bool → switch conversion and the allow-list.
func TestRealHostScript(t *testing.T) {
	exe := findExe()
	if exe == "" {
		t.Skip("PowerShell 7 is not installed")
	}
	h, err := Start(exe, []string{"Write-Output", "Out-Null"})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()

	out, err := h.Invoke(ctx, "Write-Output", map[string]any{"InputObject": "x'; Remove-Item -Recurse /; '", "NoEnumerate": true})
	if err != nil {
		t.Fatal(err)
	}
	var s string
	if len(out) != 1 || json.Unmarshal(out[0], &s) != nil || s != "x'; Remove-Item -Recurse /; '" {
		t.Fatalf("out %s", out)
	}

	_, err = h.Invoke(ctx, "Remove-Item", map[string]any{"Path": "nothing"})
	var pe *Error
	if !errors.As(err, &pe) || !strings.Contains(pe.Message, "not allowed") {
		t.Fatalf("refused cmdlet: %v", err)
	}
	// Non-ASCII survives both directions (Russian display names, rules).
	out, err = h.Invoke(ctx, "Write-Output", map[string]any{"InputObject": "Правило «Ёлка» — ü"})
	if err != nil || json.Unmarshal(out[0], &s) != nil || s != "Правило «Ёлка» — ü" {
		t.Fatalf("non-ASCII round trip: %s %v", out, err)
	}
	// A cmdlet without output answers an empty list.
	if out, err = h.Invoke(ctx, "Out-Null", map[string]any{"InputObject": 1}); err != nil || len(out) != 0 {
		t.Fatalf("no output: %s %v", out, err)
	}
	// The host survives a refused call.
	if _, err := h.Invoke(ctx, "Write-Output", map[string]any{"InputObject": 1}); err != nil {
		t.Fatalf("after refusal: %v", err)
	}
}

// TestRealHostRunsOnlyTrustedScripts: a pack script runs when its hash was
// trusted at start, any other script is refused, parameters stay data.
func TestRealHostRunsOnlyTrustedScripts(t *testing.T) {
	exe := findExe()
	if exe == "" {
		t.Skip("PowerShell 7 is not installed")
	}
	script := "param($Mode, $Inputs)\n[pscustomobject]@{ mode = $Mode; name = $Inputs.name }"
	h, err := StartWithScripts(exe, nil, []string{ScriptHash(script)})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	out, err := h.RunScript(ctx, script, map[string]any{"Mode": "read", "Inputs": map[string]any{"name": "x'; Remove-Item /; '"}})
	if err != nil {
		t.Fatal(err)
	}
	var row struct{ Mode, Name string }
	if len(out) != 1 || json.Unmarshal(out[0], &row) != nil || row.Mode != "read" || row.Name != "x'; Remove-Item /; '" {
		t.Fatalf("out %s", out)
	}
	if _, err := h.RunScript(ctx, script+"\n# changed", nil); err == nil || !strings.Contains(err.Error(), "not trusted") {
		t.Fatalf("an untrusted script must be refused: %v", err)
	}
}
