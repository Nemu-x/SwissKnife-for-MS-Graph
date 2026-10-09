package actions

import "testing"

// Every built-in action has a published capability name, and no name is
// used twice.
func TestEveryActionHasACapability(t *testing.T) {
	seen := map[string]string{}
	for _, a := range Builtin() {
		if a.Capability == "" {
			t.Errorf("%s has no capability name", a.ID)
		}
		if o, dup := seen[a.Capability]; dup {
			t.Errorf("%s and %s share capability %s", a.ID, o, a.Capability)
		}
		seen[a.Capability] = a.ID
	}
	if len(seen) != len(capabilities) {
		t.Errorf("capabilities lists %d names for %d actions", len(capabilities), len(seen))
	}
}
