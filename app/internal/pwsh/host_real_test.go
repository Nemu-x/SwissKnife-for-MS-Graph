package pwsh

import (
	"context"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
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
	// ConstrainedLanguage: objects are hashtables ([pscustomobject] is .NET).
	script := "param($Mode, $Inputs)\n@{ mode = $Mode; name = $Inputs.name }"
	h, err := StartWithScripts(exe, nil, []TrustedScript{{Hash: ScriptHash(script)}})
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

// TestRealHostSandboxesPackScripts: a pack script calls only what its pack
// declared, by name, and runs in ConstrainedLanguage.
func TestRealHostSandboxesPackScripts(t *testing.T) {
	exe := findExe()
	if exe == "" {
		t.Skip("PowerShell 7 is not installed")
	}
	cases := map[string]string{ // script → refusal ("" = runs)
		"Get-Date | Out-Null; Write-Output (Get-Random -Maximum 1)":  "",
		"function Helper { 'ok' }; Helper | ForEach-Object { $_ }":   "",
		"Remove-Item -Path nothing":                                  "did not declare the command Remove-Item",
		"$c = 'Get-Random'; & $c":                                    "must name the commands",
		"& ('Get-' + 'Random')":                                      "must name the commands",
		"Invoke-Expression 'Get-Random'":                             "cannot call Invoke-Expression",
		"[System.IO.File]::Exists('x')":                              "Method invocation is supported only",
		"1..2 | ForEach-Object -Parallel { [IO.File]::Exists('x') }": "cannot use -Parallel",
		"$ExecutionContext.InvokeCommand.InvokeScript('Get-Random')": "cannot use $ExecutionContext",
		"${function:Get-Random}":                                     "cannot read",
		"Get-Random -AsJob":                                          "cannot use -AsJob",
		". { Get-Random }":                                           "",
		"$global:x = 1":                                              "cannot use $global:x",
		"$script:allow = 1":                                          "cannot use $script:allow",
		"'a' | ForEach-Object GetType":                               "a script block",
		"$n = 'GetType'; 'a' | % $n":                                 "a script block",
		"'a' | ForEach-Object -MemberName GetType":                   "cannot use ForEach-Object -MemberName",
		"@('a').ForEach('GetType')":                                  "give .ForEach() a script block",
		"$m = 'ToString'; 'a'.$m()":                                  "must name the methods",
		"1..3 | ForEach-Object -Process { $_ } -ErrorAction Stop":    "",
		"1..3 | % { $_ } | Where-Object { $_ -gt 1 }":                "",
		"function global:Reply { }":                                  "cannot define global:Reply",
		// The host's state is out of reach (dynamic scoping would expose it).
		"if ($null -ne $packSafe -or $null -ne $scripts) { throw 'host state visible' }": "",
	}
	var trusted []TrustedScript
	for sc := range cases {
		trusted = append(trusted, TrustedScript{Hash: ScriptHash(sc), Cmdlets: []string{"Get-Random"}})
	}
	h, err := StartWithScripts(exe, nil, trusted)
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	for sc, want := range cases {
		_, err := h.RunScript(ctx, sc, nil)
		switch {
		case want == "" && err != nil && sc != ". { Get-Random }":
			t.Errorf("%q: %v", sc, err)
		case want != "" && (err == nil || !strings.Contains(err.Error(), want)):
			t.Errorf("%q: got %v, want %q", sc, err, want)
		}
	}
	// The host itself is still in FullLanguage after a constrained script.
	if _, err := h.RunScript(ctx, "Get-Date | Out-Null; Write-Output (Get-Random -Maximum 1)", nil); err != nil {
		t.Fatalf("after the sandboxed runs: %v", err)
	}
}

// The sample pack's scripts pass the sandbox: run without Exchange, they get
// as far as calling Get-Mailbox.
func TestRealHostAcceptsTheSamplePack(t *testing.T) {
	exe := findExe()
	if exe == "" {
		t.Skip("PowerShell 7 is not installed")
	}
	var scripts []string
	var trusted []TrustedScript
	for _, f := range []string{"list.ps1", "set.ps1"} {
		b, err := os.ReadFile(filepath.Join("..", "..", "..", "packs", "litigation-hold", f))
		if err != nil {
			t.Fatal(err)
		}
		scripts = append(scripts, string(b))
		trusted = append(trusted, TrustedScript{Hash: ScriptHash(string(b)), Cmdlets: []string{"Get-Mailbox", "Set-Mailbox"}})
	}
	h, err := StartWithScripts(exe, nil, trusted)
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	for _, sc := range scripts {
		_, err := h.RunScript(ctx, sc, map[string]any{"Mode": "plan", "Inputs": map[string]any{"mailbox": "a@b.c", "state": "on"}})
		if err == nil || !strings.Contains(err.Error(), "Get-Mailbox") || strings.Contains(err.Error(), "pack") {
			t.Fatalf("the sample should reach Get-Mailbox: %v", err)
		}
	}
}
