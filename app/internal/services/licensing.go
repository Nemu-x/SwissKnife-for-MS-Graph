package services

import (
	"encoding/json"
	"net/url"
	"strconv"

	"swissknife-app/internal/session"
)

type LicensingService struct {
	s *session.Session
}

func NewLicensingService(s *session.Session) *LicensingService { return &LicensingService{s: s} }

func (l *LicensingService) Skus() ([]json.RawMessage, error) {
	c, err := l.s.Client()
	if err != nil {
		return nil, err
	}
	return c.ListAll(l.s.Ctx(), "/subscribedSkus", nil, 0)
}

func (l *LicensingService) Assign(user string, addSkuIDs, removeSkuIDs []string) (json.RawMessage, error) {
	if err := l.s.GuardWrite(); err != nil {
		return nil, err
	}
	c, err := l.s.Client()
	if err != nil {
		return nil, err
	}
	add := make([]map[string]any, 0, len(addSkuIDs))
	for _, id := range addSkuIDs {
		add = append(add, map[string]any{"skuId": id})
	}
	if removeSkuIDs == nil {
		removeSkuIDs = []string{}
	}
	body := map[string]any{"addLicenses": add, "removeLicenses": removeSkuIDs}
	var out json.RawMessage
	err = c.Post(l.s.Ctx(), "/users/"+url.PathEscape(user)+"/assignLicense", body, &out)
	l.s.Record("licensing.assign", user,
		"add="+strconv.Itoa(len(addSkuIDs))+" remove="+strconv.Itoa(len(removeSkuIDs)), err)
	return out, err
}
