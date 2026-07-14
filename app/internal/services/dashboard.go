package services

import (
	"encoding/json"

	"swissknife-app/internal/session"
)

// DashboardService aggregates a tenant overview for the landing screen.
type DashboardService struct {
	s *session.Session
}

func NewDashboardService(s *session.Session) *DashboardService { return &DashboardService{s: s} }

// LicenseLine is a per-SKU consumption summary.
type LicenseLine struct {
	SkuPartNumber string `json:"skuPartNumber"`
	Consumed      int    `json:"consumed"`
	Total         int    `json:"total"`
}

// DashboardSummary is the overview shown after connecting.
type DashboardSummary struct {
	OrgName       string        `json:"orgName"`
	Users         int           `json:"users"`
	Groups        int           `json:"groups"`
	Domains       int           `json:"domains"`
	LicensesUsed  int           `json:"licensesUsed"`
	LicensesTotal int           `json:"licensesTotal"`
	Licenses      []LicenseLine `json:"licenses"`
}

// Summary gathers counts and license usage. Individual failures degrade to zero
// values so a missing permission does not blank the whole dashboard.
func (d *DashboardService) Summary() (*DashboardSummary, error) {
	c, err := d.s.Client()
	if err != nil {
		return nil, err
	}
	ctx := d.s.Ctx()
	out := &DashboardSummary{}

	// organization name
	var org struct {
		Value []struct {
			DisplayName string `json:"displayName"`
		} `json:"value"`
	}
	if err := c.Get(ctx, "/organization", nil, &org); err == nil && len(org.Value) > 0 {
		out.OrgName = org.Value[0].DisplayName
	}

	out.Users, _ = c.Count(ctx, "users")
	out.Groups, _ = c.Count(ctx, "groups")

	// domains
	if domains, err := c.ListAll(ctx, "/domains", nil, 0); err == nil {
		out.Domains = len(domains)
	}

	// licenses
	if skus, err := c.ListAll(ctx, "/subscribedSkus", nil, 0); err == nil {
		for _, raw := range skus {
			var sku struct {
				SkuPartNumber string `json:"skuPartNumber"`
				ConsumedUnits int    `json:"consumedUnits"`
				PrepaidUnits  struct {
					Enabled int `json:"enabled"`
				} `json:"prepaidUnits"`
				CapabilityStatus string `json:"capabilityStatus"`
			}
			if json.Unmarshal(raw, &sku) != nil {
				continue
			}
			out.Licenses = append(out.Licenses, LicenseLine{
				SkuPartNumber: sku.SkuPartNumber,
				Consumed:      sku.ConsumedUnits,
				Total:         sku.PrepaidUnits.Enabled,
			})
			out.LicensesUsed += sku.ConsumedUnits
			out.LicensesTotal += sku.PrepaidUnits.Enabled
		}
	}

	return out, nil
}
