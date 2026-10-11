package cli

import (
	"bytes"
	"errors"
	"strings"
	"testing"

	"swissknife-app/internal/engine"
)

var secretCatalog = []engine.CatalogEntry{{Manifest: engine.Manifest{ID: "x.reset", Fields: []engine.Field{
	{Name: "user", Kind: engine.FieldText, Required: true},
	{Name: "newPassword", Kind: engine.FieldSecret, Required: true},
	{Name: "recovery", Kind: engine.FieldSecret},
}}}}

func TestActionSecretOnCommandLineRefused(t *testing.T) {
	f := newFake(t, one)
	code := run(f.e, []string{"action", "ad.resetPassword", "adUser=ann", "newPassword=hunter2"})
	if code != exitUsage {
		t.Fatalf("exit %d, want %d", code, exitUsage)
	}
	if !strings.Contains(f.err.String(), "newPassword is a secret") || strings.Contains(f.err.String(), "hunter2") {
		t.Fatalf("stderr %q", f.err.String())
	}
}

func TestReadSecretsFromStdin(t *testing.T) {
	var stderr bytes.Buffer
	e := &env{stdin: strings.NewReader("s3cret pw\r\nsecond\n"), stderr: &stderr}

	// Missing required secret and "-" both read the next stdin line.
	in := engine.Inputs{"user": "ann", "recovery": "-"}
	if err := readSecrets(e, secretCatalog, "x.reset", in); err != nil {
		t.Fatal(err)
	}
	if in["newPassword"] != "s3cret pw" || in["recovery"] != "second" {
		t.Fatalf("inputs %v", in)
	}
	if stderr.Len() != 0 {
		t.Fatalf("a piped secret must not prompt: %q", stderr.String())
	}

	// An optional secret that was not asked for stays unset.
	in = engine.Inputs{"user": "ann"}
	e.stdin = strings.NewReader("only\n")
	if err := readSecrets(e, secretCatalog, "x.reset", in); err != nil {
		t.Fatal(err)
	}
	if _, ok := in["recovery"]; ok || in["newPassword"] != "only" {
		t.Fatalf("inputs %v", in)
	}

	// A last line without a newline still counts; empty stdin is refused.
	in = engine.Inputs{}
	e.stdin = strings.NewReader("noeol")
	if err := readSecrets(e, secretCatalog, "x.reset", in); err != nil || in["newPassword"] != "noeol" {
		t.Fatalf("err %v inputs %v", err, in)
	}
	e.stdin = strings.NewReader("")
	err := readSecrets(e, secretCatalog, "x.reset", engine.Inputs{})
	var ue usageError
	if !errors.As(err, &ue) {
		t.Fatalf("empty stdin: %v", err)
	}

	// Other actions are untouched.
	in = engine.Inputs{"newPassword": "x"}
	if err := readSecrets(e, secretCatalog, "other", in); err != nil || in["newPassword"] != "x" {
		t.Fatalf("err %v inputs %v", err, in)
	}
}
