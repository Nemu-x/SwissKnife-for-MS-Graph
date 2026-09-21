package graphapi

import (
	"context"
	"errors"
	"time"
)

// importSessionMargin renews an import session this long before Graph says it
// expires, so a batch never starts against a URL that dies mid-way.
const importSessionMargin = 2 * time.Minute

// ImportSession is an opaque handle on a Graph mailbox import session. Graph
// answers createImportSession with a pre-authorized upload URL whose token
// rides in the query string; that URL never leaves this package (ADR-002:
// services do not see raw credentials). The session is created lazily, renewed
// ahead of its expiry, and once more when Graph answers 401 to an upload.
type ImportSession struct {
	c          *Client
	createPath string // Graph path of the createImportSession action
	url        string
	expires    time.Time
}

// NewImportSession prepares a session for the given createImportSession path
// (e.g. /admin/exchange/mailboxes/{id}/createImportSession). No request is
// made until the first Post.
func (c *Client) NewImportSession(createPath string) *ImportSession {
	return &ImportSession{c: c, createPath: createPath}
}

func (s *ImportSession) ensure(ctx context.Context) error {
	if s.url != "" && time.Until(s.expires) > importSessionMargin {
		return nil
	}
	var out struct {
		ImportURL  string    `json:"importUrl"`
		Expiration time.Time `json:"expirationDateTime"`
	}
	if err := s.c.Post(ctx, s.createPath, nil, &out); err != nil {
		return err
	}
	if out.ImportURL == "" {
		return errors.New("graph: createImportSession returned no importUrl")
	}
	s.url = out.ImportURL
	s.expires = out.Expiration
	if s.expires.IsZero() {
		s.expires = time.Now().Add(45 * time.Minute)
	}
	return nil
}

// Post uploads one JSON body through the session and decodes the response
// into out. A 401 (the session token died under us) takes a fresh session and
// retries the upload once.
func (s *ImportSession) Post(ctx context.Context, body, out any) error {
	if err := s.ensure(ctx); err != nil {
		return err
	}
	err := s.c.PostPreauthorized(ctx, s.url, body, out)
	var ge *GraphError
	if errors.As(err, &ge) && ge.StatusCode == 401 {
		s.url = ""
		if rerr := s.ensure(ctx); rerr != nil {
			return rerr
		}
		err = s.c.PostPreauthorized(ctx, s.url, body, out)
	}
	return err
}
