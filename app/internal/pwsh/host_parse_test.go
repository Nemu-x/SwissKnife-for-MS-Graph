package pwsh

import (
	"bytes"
	"context"
	"encoding/json"
	"regexp"
	"testing"
	"time"
)

// The request line (connect requests carry access tokens) must never be the
// input of a cmdlet: Module Logging writes cmdlet bindings to event 4103.
func TestHostScriptDoesNotParseRequestsWithACmdlet(t *testing.T) {
	if regexp.MustCompile(`(?i)\$line\s*\|\s*ConvertFrom-Json|ConvertFrom-Json\s+.*\$line`).Match(hostScript) {
		t.Fatal("host.ps1 pipes the request line into ConvertFrom-Json")
	}
}

// TestRealHostDecodesNestedRequests: the .NET request parser yields the same
// shapes ConvertFrom-Json -AsHashtable did (nested objects, arrays, numbers,
// bools, null) so splatting keeps working.
func TestRealHostDecodesNestedRequests(t *testing.T) {
	exe := findExe()
	if exe == "" {
		t.Skip("PowerShell 7 is not installed")
	}
	h, err := Start(exe, []string{"Write-Output"})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()

	in := map[string]any{
		"s": "x", "i": 42, "big": 9007199254740993, "f": 1.5, "t": true, "n": nil,
		"arr": []any{"a", 2, []any{}}, "one": []any{"only"}, "empty": []any{},
		"obj": map[string]any{"k": map[string]any{"deep": "v"}},
	}
	out, err := h.Invoke(ctx, "Write-Output", map[string]any{"InputObject": in, "NoEnumerate": true})
	if err != nil {
		t.Fatal(err)
	}
	if len(out) != 1 {
		t.Fatalf("out %s", out)
	}
	var got map[string]any
	d := json.NewDecoder(bytes.NewReader(out[0]))
	d.UseNumber()
	if err := d.Decode(&got); err != nil {
		t.Fatal(err)
	}
	checks := map[string]string{
		"s": `"x"`, "i": `42`, "big": `9007199254740993`, "f": `1.5`, "t": `true`, "n": `null`,
		"arr": `["a",2,[]]`, "one": `["only"]`, "empty": `[]`, "obj": `{"k":{"deep":"v"}}`,
	}
	for k, want := range checks {
		b, _ := json.Marshal(got[k])
		if string(b) != want {
			t.Errorf("%s = %s, want %s", k, b, want)
		}
	}
	// A malformed line is answered with an error and the host keeps going.
	if _, err := h.call(ctx, map[string]any{"op": "nope"}); err == nil {
		t.Fatal("unknown op accepted")
	}
	if _, err := h.Invoke(ctx, "Write-Output", map[string]any{"InputObject": 1}); err != nil {
		t.Fatalf("after an error: %v", err)
	}
}
