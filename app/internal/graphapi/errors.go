package graphapi

import "fmt"

// GraphError — структурированная ошибка Microsoft Graph.
// UI показывает Code/Message/RequestID вместо сырого дампа (ADR-002).
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

// IsNotFound — 404 от Graph.
func IsNotFound(err error) bool {
	ge, ok := err.(*GraphError)
	return ok && ge.StatusCode == 404
}

// IsForbidden — 403: не хватает permissions (UI подсказывает, каких — ADR-002).
func IsForbidden(err error) bool {
	ge, ok := err.(*GraphError)
	return ok && ge.StatusCode == 403
}
