package graphapi

import (
	"errors"
	"fmt"
)

// GraphError is a structured Microsoft Graph error.
// The UI shows Code/Message/RequestID instead of a raw dump (ADR-002).
type GraphError struct {
	StatusCode int
	Code       string
	Message    string
	RequestID  string
	Path       string // request URL path — lets callers derive permission hints
}

func (e *GraphError) Error() string {
	if e.RequestID != "" {
		return fmt.Sprintf("graph: %d %s: %s (requestId=%s)", e.StatusCode, e.Code, e.Message, e.RequestID)
	}
	return fmt.Sprintf("graph: %d %s: %s", e.StatusCode, e.Code, e.Message)
}

// IsNotFound reports a 404 from Graph (errors.As — wrapped errors included).
func IsNotFound(err error) bool {
	var ge *GraphError
	return errors.As(err, &ge) && ge.StatusCode == 404
}

// IsForbidden reports a 403: missing permissions (the UI hints which — ADR-002).
func IsForbidden(err error) bool {
	var ge *GraphError
	return errors.As(err, &ge) && ge.StatusCode == 403
}
