package graphapi

import (
	"context"
	"errors"
	"net/url"
)

// InitialDomain returns the tenant's initial <name>.onmicrosoft.com domain —
// what Exchange expects as the organization (Connect-ExchangeOnline
// -Organization, the app-only X-AnchorMailbox routing key).
func InitialDomain(ctx context.Context, c *Client) (string, error) {
	var org struct {
		Value []struct {
			Domains []struct {
				Name      string `json:"name"`
				IsInitial bool   `json:"isInitial"`
			} `json:"verifiedDomains"`
		} `json:"value"`
	}
	if err := c.Get(ctx, "/organization", url.Values{"$select": {"verifiedDomains"}}, &org); err != nil {
		return "", err
	}
	for _, o := range org.Value {
		for _, d := range o.Domains {
			if d.IsInitial {
				return d.Name, nil
			}
		}
	}
	return "", errors.New("the tenant's initial .onmicrosoft.com domain was not found")
}
