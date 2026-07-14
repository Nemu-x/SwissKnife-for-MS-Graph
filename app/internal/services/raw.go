package services

import (
	"encoding/json"
	"errors"
	"net/http"
	"strings"

	"swissknife-app/internal/session"
)

// RawService is the Graph playground: arbitrary requests.
// History/favorites live in the frontend (webview localStorage).
type RawService struct {
	s *session.Session
}

func NewRawService(s *session.Session) *RawService { return &RawService{s: s} }

var allowedMethods = map[string]bool{
	http.MethodGet: true, http.MethodPost: true, http.MethodPatch: true,
	http.MethodPut: true, http.MethodDelete: true,
}

// Send executes a request. body is a JSON string (empty = no body).
func (r *RawService) Send(method, path, body string) (json.RawMessage, error) {
	method = strings.ToUpper(strings.TrimSpace(method))
	if !allowedMethods[method] {
		return nil, errors.New("method must be GET/POST/PATCH/PUT/DELETE")
	}
	if method != http.MethodGet {
		if err := r.s.GuardWrite(); err != nil {
			return nil, err
		}
	}
	c, err := r.s.Client()
	if err != nil {
		return nil, err
	}

	var payload any
	if strings.TrimSpace(body) != "" {
		var v json.RawMessage
		if err := json.Unmarshal([]byte(body), &v); err != nil {
			return nil, errors.New("request body is not valid JSON: " + err.Error())
		}
		payload = v
	}

	var out json.RawMessage
	err = c.Do(r.s.Ctx(), method, path, nil, payload, &out)
	if method != http.MethodGet {
		r.s.Record("raw."+strings.ToLower(method), path, "", err)
	}
	if err != nil {
		return nil, err
	}
	if out == nil {
		out = json.RawMessage(`{"status":"no content"}`)
	}
	return out, nil
}
