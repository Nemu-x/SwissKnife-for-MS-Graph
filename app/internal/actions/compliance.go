package actions

import (
	"context"
	"errors"
	"fmt"
	"regexp"
	"strings"
	"time"

	"swissknife-app/internal/engine"
)

// Phishing clean-up through Microsoft Graph eDiscovery: a search over every
// mailbox, an estimate, then purgeData. Graph supports this for both sign-in
// kinds, on every platform — no Security & Compliance PowerShell needed.
//
// Unlike other plans, this preview writes something: the eDiscovery case and
// search the estimate needs. They hold no tenant data change, and stay as the
// audit trail of the purge (named after the sender and the date).

// pollEvery and the deadlines are variables so tests do not wait.
var (
	pollEvery        = 3 * time.Second
	estimateDeadline = 5 * time.Minute
	purgeDeadline    = 10 * time.Minute
	purgeRounds      = 5 // Microsoft purges at most 100 items per mailbox per run
)

func complianceActions() []engine.Action {
	return []engine.Action{{
		Manifest: engine.Manifest{
			ID: "mail.purge", Page: "audit", Danger: engine.Destructive,
			Fields: []engine.Field{
				{Name: "sender", Kind: engine.FieldText, Required: true},
				{Name: "subject", Kind: engine.FieldText},
				{Name: "since", Kind: engine.FieldText},
				{Name: "purgeType", Kind: engine.FieldChoice, Required: true, Options: []string{"recoverable", "permanent"}, Default: "recoverable"},
			},
			ConfirmField: "sender",
			Permissions:  []string{"eDiscovery.ReadWrite.All", "Purview: Organization Management (Search And Purge)"},
		},
		Impls: []engine.Impl{graphPurge{}},
	}}
}

var (
	senderRe = regexp.MustCompile(`^([^@\s"()]+@)?[A-Za-z0-9.-]+\.[A-Za-z]{2,}$`)
	dateRe   = regexp.MustCompile(`^\d{4}-\d{2}-\d{2}$`)
)

// purgeQuery builds the KQL. Values are validated or stripped of the
// characters KQL would read as syntax, so the operator cannot widen the
// search by accident (a stray quote could otherwise match every mailbox).
func purgeQuery(in engine.Inputs) (string, error) {
	sender := strings.TrimSpace(in["sender"])
	if !senderRe.MatchString(sender) {
		return "", errors.New("sender must be an address (bad@example.com) or a domain (example.com)")
	}
	q := `from:"` + sender + `"`
	if subj := strings.Map(func(r rune) rune {
		if strings.ContainsRune(`"()\:*`, r) {
			return ' '
		}
		return r
	}, strings.TrimSpace(in["subject"])); strings.TrimSpace(subj) != "" {
		q += ` AND subject:"` + strings.Join(strings.Fields(subj), " ") + `"`
	}
	if since := strings.TrimSpace(in["since"]); since != "" {
		if !dateRe.MatchString(since) {
			return "", errors.New("received since must be a date like 2026-10-01")
		}
		q += " AND received>=" + since
	}
	return q, nil
}

type graphPurge struct{}

func (graphPurge) Backend() engine.Backend { return engine.BackendGraph }

type estimate struct {
	Status           string `json:"status"`
	IndexedItemCount int64  `json:"indexedItemCount"`
	IndexedItemsSize int64  `json:"indexedItemsSize"`
	MailboxCount     int    `json:"mailboxCount"`
}

// waitOperation polls an eDiscovery operation URL until it finishes.
func waitOperation(ctx context.Context, env engine.Env, loc string, deadline time.Duration, into any) error {
	ctx, cancel := context.WithTimeout(ctx, deadline)
	defer cancel()
	for {
		var op struct {
			Status string `json:"status"`
		}
		if err := env.Graph.Get(ctx, loc, nil, &op); err != nil {
			return err
		}
		switch op.Status {
		case "succeeded", "partiallySucceeded":
			if into != nil {
				return env.Graph.Get(ctx, loc, nil, into)
			}
			return nil
		case "failed", "submissionFailed":
			return fmt.Errorf("the eDiscovery operation %s", op.Status)
		}
		select {
		case <-ctx.Done():
			return fmt.Errorf("the eDiscovery operation did not finish in time: %w", ctx.Err())
		case <-time.After(pollEvery):
		}
	}
}

func runEstimate(env engine.Env, searchPath string) (*estimate, error) {
	loc, err := env.Graph.PostForLocation(env.Ctx, searchPath+"/estimateStatistics", nil, map[string]any{})
	if err != nil {
		return nil, err
	}
	if loc == "" {
		return nil, errors.New("the estimate did not return an operation to follow")
	}
	var est estimate
	if err := waitOperation(env.Ctx, env, loc, estimateDeadline, &est); err != nil {
		return nil, err
	}
	return &est, nil
}

func (graphPurge) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	query, err := purgeQuery(in)
	if err != nil {
		return nil, err
	}
	name := fmt.Sprintf("SwissKnife purge %s %s", strings.TrimSpace(in["sender"]), time.Now().UTC().Format("2006-01-02 15:04"))
	var cs struct {
		ID string `json:"id"`
	}
	if err := env.Graph.Post(env.Ctx, "/security/cases/ediscoveryCases", map[string]any{
		"displayName": name,
		"description": "Created by SwissKnife to find and remove a phishing message. Query: " + query,
	}, &cs); err != nil {
		return nil, err
	}
	var search struct {
		ID string `json:"id"`
	}
	casePath := "/security/cases/ediscoveryCases/" + cs.ID
	if err := env.Graph.Post(env.Ctx, casePath+"/searches", map[string]any{
		"displayName":      name,
		"contentQuery":     query,
		"dataSourceScopes": "allTenantMailboxes",
	}, &search); err != nil {
		return nil, err
	}
	searchPath := casePath + "/searches/" + search.ID
	est, err := runEstimate(env, searchPath)
	if err != nil {
		return nil, err
	}
	// Numbers only: the field label ("messages / mailboxes") translates.
	ch := engine.Change{
		Target: query,
		Field:  "messagesInMailboxes", Op: "remove",
		Before: fmt.Sprintf("%d / %d", est.IndexedItemCount, est.MailboxCount),
		Note:   "purge." + in["purgeType"],
		Ref:    map[string]string{"search": searchPath},
	}
	if est.IndexedItemCount == 0 {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

// Apply purges. One run removes at most 100 items per mailbox, so a
// permanent purge re-estimates and runs again until nothing matches. A
// recoverable purge runs once: soft-deleted items stay in Recoverable Items,
// where the search still finds them, so the count would never reach zero.
func (graphPurge) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	purgeType, rounds := "recoverable", 1
	if in["purgeType"] == "permanent" {
		purgeType, rounds = "permanentlyDelete", purgeRounds
	}
	path := ch.Ref["search"]
	for round := 0; round < rounds; round++ {
		loc, err := env.Graph.PostForLocation(env.Ctx, path+"/purgeData", nil,
			map[string]any{"purgeType": purgeType, "purgeAreas": "mailboxes"})
		if err != nil {
			return err
		}
		if loc != "" {
			if err := waitOperation(env.Ctx, env, loc, purgeDeadline, nil); err != nil {
				return err
			}
		}
		if rounds == 1 {
			return nil
		}
		est, err := runEstimate(env, path)
		if err != nil {
			return err
		}
		if est.IndexedItemCount == 0 {
			return nil
		}
	}
	return fmt.Errorf("messages still match after %d purge runs — run the action again", rounds)
}
