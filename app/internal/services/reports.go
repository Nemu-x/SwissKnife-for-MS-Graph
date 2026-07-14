package services

import (
	"errors"
	"net/url"

	"swissknife-app/internal/session"
)

// ReportsService — Microsoft 365 usage reports (returned as CSV by Graph).
type ReportsService struct {
	s *session.Session
}

func NewReportsService(s *session.Session) *ReportsService { return &ReportsService{s: s} }

// report name -> Graph function. period is D7|D30|D90|D180.
var reportFuncs = map[string]string{
	"office365ActiveUsers": "getOffice365ActiveUserDetail",
	"oneDriveUsage":        "getOneDriveUsageAccountDetail",
	"mailboxUsage":         "getMailboxUsageDetail",
	"teamsUserActivity":    "getTeamsUserActivityUserDetail",
	"sharePointUsage":      "getSharePointSiteUsageDetail",
}

// Names returns the available report identifiers.
func (r *ReportsService) Names() []string {
	names := make([]string, 0, len(reportFuncs))
	for k := range reportFuncs {
		names = append(names, k)
	}
	return names
}

// CSV fetches a usage report as CSV text. The UI offers it for download.
func (r *ReportsService) CSV(report, period string) (string, error) {
	fn, ok := reportFuncs[report]
	if !ok {
		return "", errors.New("unknown report: " + report)
	}
	if period == "" {
		period = "D30"
	}
	c, err := r.s.Client()
	if err != nil {
		return "", err
	}
	path := "/reports/" + fn + "(period='" + url.QueryEscape(period) + "')"
	return c.GetText(r.s.Ctx(), path, nil)
}
