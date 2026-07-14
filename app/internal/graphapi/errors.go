package graphapi

import "fmt"

// GraphError is a structured Microsoft Graph error.
// The UI shows Code/Message/RequestID instead of a raw dump (ADR-002).
type GraphError struct {
	StatusCode int
	Code       string
	Message    string
	RequestID  string
}

func (e *GraphError) Error() string {
	if e.RequestID != "" {
		return fmt.Sprintf("graph: %d %s: %s (requestId=%s)", e.StatusCode, e.Code, e.Message, e.RequestID)
	}
	return fmt.Sprintf("graph: %d %s: %s", e.StatusCode, e.Code, e.Message)
}

// IsNotFound reports a 404 from Graph.
func IsNotFound(err error) bool {
	ge, ok := err.(*GraphError)
	return ok && ge.StatusCode == 404
}

// IsForbidden reports a 403: missing permissions (the UI hints which — ADR-002).
func IsForbidden(err error) bool {
	ge, ok := err.(*GraphError)
	return ok && ge.StatusCode == 403
}
