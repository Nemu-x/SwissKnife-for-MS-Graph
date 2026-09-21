//go:build !windows

package cli

// attachConsole is a no-op outside Windows: console subsystems inherit the
// terminal's std handles as usual.
func attachConsole() {}
