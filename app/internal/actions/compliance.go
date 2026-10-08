package actions

import (
	"context"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"regexp"
	"strings"
	"time"

	"swissknife-app/internal/engine"
	"swissknife-app/internal/session"
)

// Phishing clean-up through Microsoft Graph eDiscovery: a search over every
// mailbox, an estimate, then purgeData. Graph supports this for both sign-in
// kinds, on every platform — no Security & Compliance PowerShell needed.
//
// This is the most dangerous action in the app — a wrong query deletes mail
// across the whole tenant — so it is fenced more than the others:
//   - the query is built only from validated values, and a bare domain needs
//     a subject or a date and may not be one of the tenant's own domains or a
//     large free-mail provider;
//   - the confirmation is the sender plus the number of messages;
//   - before each purge the search is read back and must still hold exactly
//     the previewed query over all mailboxes;
//   - partial results count as failures.
//
// Unlike other plans, this preview writes: the eDiscovery case and search the
// estimate needs. It therefore respects read-only mode, removes the case
// again if the preview fails, and otherwise leaves it in Purview as the
// record of the clean-up.

// pollEvery and the deadlines are variables so tests do not wait.
var (
	pollEvery        = 3 * time.Second
	estimateDeadline = 5 * time.Minute
	purgeDeadline    = 10 * time.Minute
	purgeRounds      = 5 // Microsoft purges at most 100 items per mailbox per run
	// settle lets the index catch up after a hard delete before the count is
	// compared again; without it, just-deleted items can still be counted.
	settle = 30 * time.Second
)

// perRunLimit is Microsoft's cap on items purged per mailbox per run.
const perRunLimit = 100

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
			Permissions:  []string{"eDiscovery.ReadWrite.All", "Domain.Read.All", "Purview: Organization Management (Search And Purge)"},
		},
		Impls: []engine.Impl{graphPurge{}},
	}}
}

var (
	// A plain address or domain: no wildcards, quotes, colons or spaces.
	senderAddrRe = regexp.MustCompile(`^[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$`)
	senderDomRe  = regexp.MustCompile(`^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$`)
	// Subject: letters, digits, spaces and plain punctuation only — anything
	// KQL could read as syntax (quotes of any kind, parentheses, colons,
	// wildcards, backslashes) is refused rather than stripped.
	subjectRe = regexp.MustCompile(`^[\p{L}\p{N} .,!?#&@$%+=_/'-]+$`)
)

// freeMail are providers whose whole domain must never be purged.
var freeMail = map[string]bool{
	"gmail.com": true, "googlemail.com": true, "outlook.com": true, "hotmail.com": true, "live.com": true,
	"msn.com": true, "yahoo.com": true, "icloud.com": true, "me.com": true, "aol.com": true,
	"proton.me": true, "protonmail.com": true, "gmx.com": true, "gmx.de": true, "web.de": true,
	"mail.ru": true, "yandex.ru": true, "ya.ru": true, "bk.ru": true, "inbox.ru": true, "list.ru": true,
	"rambler.ru": true, "zoho.com": true, "qq.com": true, "163.com": true, "onmicrosoft.com": true,
}

// purgeQuery builds the KQL from validated values. internal holds the
// tenant's own domains.
func purgeQuery(in engine.Inputs, internal map[string]bool) (string, error) {
	sender := strings.ToLower(strings.TrimSpace(in["sender"]))
	subject := strings.Join(strings.Fields(in["subject"]), " ")
	since := strings.TrimSpace(in["since"])

	bareDomain := false
	switch {
	case senderAddrRe.MatchString(sender):
		// A colleague's mailbox (a compromised account sending phishing) is a
		// valid target, but all of its mail is not: narrow it.
		if internal[sender[strings.LastIndex(sender, "@")+1:]] && subject == "" && since == "" {
			return "", errors.New("an address in this tenant needs a subject or a date as well, to keep the search narrow")
		}
	case senderDomRe.MatchString(sender):
		bareDomain = true
	default:
		return "", errors.New("sender must be an address (bad@example.com) or a domain (example.com) — no wildcards or quotes")
	}
	if bareDomain {
		if internal[sender] {
			return "", errors.New("refusing to purge a whole domain of this tenant — enter the sender's full address")
		}
		if freeMail[sender] {
			return "", errors.New("refusing to purge a whole free-mail domain — enter the sender's full address")
		}
		if subject == "" && since == "" {
			return "", errors.New("a whole domain needs a subject or a date as well, to keep the search narrow")
		}
	}

	q := `from:"` + sender + `"`
	if subject != "" {
		if !subjectRe.MatchString(subject) {
			return "", errors.New("the subject may contain letters, digits, spaces and plain punctuation only (no quotes, brackets, colons or wildcards)")
		}
		q += ` AND subject:"` + subject + `"`
	}
	if since != "" {
		d, err := time.Parse("2006-01-02", since)
		if err != nil || d.After(time.Now()) {
			return "", errors.New("received since must be a past date like 2026-10-01")
		}
		q += " AND received>=" + d.Format("2006-01-02")
	}
	return q, nil
}

type graphPurge struct{}

func (graphPurge) Backend() engine.Backend { return engine.BackendGraph }

// operation is an eDiscovery operation: its status plus, for an estimate,
// the counts.
type operation struct {
	Status           string `json:"status"`
	IndexedItemCount int64  `json:"indexedItemCount"`
	MailboxCount     int    `json:"mailboxCount"`
}

// errOutcomeUnknown marks a purge that was submitted but whose end was not
// seen: running it again would start a second purge.
var errOutcomeUnknown = errors.New("the purge was submitted but its outcome is unknown — check the case in Microsoft Purview before running it again")

// waitOperation polls an operation URL until it finishes. A few failed polls
// in a row are tolerated (the operation keeps running on Microsoft's side).
func waitOperation(ctx context.Context, env engine.Env, loc string, deadline time.Duration) (*operation, error) {
	ctx, cancel := context.WithTimeout(ctx, deadline)
	defer cancel()
	failures := 0
	for {
		var op operation
		err := env.Graph.Get(ctx, loc, nil, &op)
		switch {
		case err != nil && ctx.Err() == nil && failures < 3:
			failures++
		case err != nil:
			return nil, err
		case op.Status == "succeeded":
			return &op, nil
		case op.Status == "partiallySucceeded":
			return &op, errors.New("partial result reported by Microsoft: some mailboxes were not processed — check the case in Microsoft Purview")
		case op.Status == "failed" || op.Status == "submissionFailed":
			return nil, fmt.Errorf("the eDiscovery operation %s", op.Status)
		default:
			failures = 0
		}
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-time.After(pollEvery):
		}
	}
}

func runEstimate(env engine.Env, searchPath string) (*operation, error) {
	loc, err := env.Graph.PostForLocation(env.Ctx, searchPath+"/estimateStatistics", nil, map[string]any{})
	if err != nil {
		return nil, err
	}
	if loc == "" {
		return nil, errors.New("the estimate did not return an operation to follow")
	}
	op, err := waitOperation(env.Ctx, env, loc, estimateDeadline)
	if err != nil {
		return nil, fmt.Errorf("estimate: %w", err)
	}
	return op, nil
}

func tenantDomains(env engine.Env) (map[string]bool, error) {
	var d struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	if err := env.Graph.Get(env.Ctx, "/domains", url.Values{"$select": {"id"}}, &d); err != nil {
		return nil, err
	}
	out := map[string]bool{}
	for _, x := range d.Value {
		out[strings.ToLower(x.ID)] = true
	}
	return out, nil
}

func (graphPurge) Plan(env engine.Env, in engine.Inputs) ([]engine.Change, error) {
	internal, err := tenantDomains(env)
	if err != nil {
		return nil, err
	}
	query, err := purgeQuery(in, internal)
	if err != nil {
		return nil, err
	}
	// The preview creates a case and a search: a write, so read-only mode
	// stops it here.
	if env.ReadOnly {
		return nil, session.ErrReadOnly
	}
	sender := strings.ToLower(strings.TrimSpace(in["sender"]))
	name := fmt.Sprintf("SwissKnife purge %s %s", sender, time.Now().UTC().Format("2006-01-02 15:04"))
	var cs struct {
		ID string `json:"id"`
	}
	if err := env.Graph.Post(env.Ctx, "/security/cases/ediscoveryCases", map[string]any{
		"displayName": name,
		"description": "Created by SwissKnife to find and remove a phishing message. Query: " + query,
	}, &cs); err != nil {
		return nil, err
	}
	casePath := "/security/cases/ediscoveryCases/" + url.PathEscape(cs.ID)
	// A failed preview leaves nothing behind.
	ok := false
	defer func() {
		if !ok {
			_ = env.Graph.Delete(context.WithoutCancel(env.Ctx), casePath)
		}
	}()

	var search struct {
		ID string `json:"id"`
	}
	if err := env.Graph.Post(env.Ctx, casePath+"/searches", map[string]any{
		"displayName":      name,
		"contentQuery":     query,
		"dataSourceScopes": "allTenantMailboxes",
	}, &search); err != nil {
		return nil, err
	}
	searchPath := casePath + "/searches/" + url.PathEscape(search.ID)
	est, err := runEstimate(env, searchPath)
	if err != nil {
		return nil, err
	}
	ok = true

	// Numbers only: the field label ("messages / mailboxes") translates. The
	// confirmation names the sender and the count, so the operator retypes
	// how much is about to go.
	note := "purge." + in["purgeType"]
	if in["purgeType"] != "permanent" && est.MailboxCount > 0 && est.IndexedItemCount > int64(perRunLimit*est.MailboxCount) {
		note = "purge.recoverableOverLimit"
	}
	ch := engine.Change{
		Target: query,
		Field:  "messagesInMailboxes", Op: "remove",
		Before: fmt.Sprintf("%d / %d", est.IndexedItemCount, est.MailboxCount),
		Note:   note,
		Ref: map[string]string{
			"search": searchPath, "query": query,
			engine.ConfirmRef: fmt.Sprintf("%s %d", sender, est.IndexedItemCount),
		},
	}
	if est.IndexedItemCount == 0 {
		ch.Op = "none"
	}
	return []engine.Change{ch}, nil
}

// verifySearch reads the search back: between preview and apply anyone with
// Purview access could have edited it, and the purge acts on what it holds
// now, not on what was previewed.
func verifySearch(env engine.Env, path, query string) error {
	var s struct {
		ContentQuery     string `json:"contentQuery"`
		DataSourceScopes string `json:"dataSourceScopes"`
	}
	if err := env.Graph.Get(env.Ctx, path, url.Values{"$select": {"contentQuery,dataSourceScopes"}}, &s); err != nil {
		return err
	}
	if s.ContentQuery != query || !strings.Contains(s.DataSourceScopes, "allTenantMailboxes") {
		return errors.New("the eDiscovery search was changed after the preview — nothing was purged; preview again")
	}
	return nil
}

// Apply purges. One run removes at most 100 items per mailbox, so a
// permanent purge re-estimates and runs again while the count keeps falling.
// A recoverable purge runs once: soft-deleted items stay in Recoverable
// Items, where the search still finds them, so the count would not fall.
func (graphPurge) Apply(env engine.Env, in engine.Inputs, ch engine.Change) error {
	purgeType, rounds := "recoverable", 1
	if in["purgeType"] == "permanent" {
		purgeType, rounds = "permanentlyDelete", purgeRounds
	}
	path, query := ch.Ref["search"], ch.Ref["query"]
	// permanentlyDelete is an evolvable enum member in v1.0.
	hdr := http.Header{"Prefer": {"include-unknown-enum-members"}}
	last := int64(-1)
	for round := 0; round < rounds; round++ {
		if err := verifySearch(env, path, query); err != nil {
			return err
		}
		loc, err := env.Graph.PostForLocationWithHeaders(env.Ctx, path+"/purgeData", nil, hdr,
			map[string]any{"purgeType": purgeType, "purgeAreas": "mailboxes"})
		if err != nil {
			return err
		}
		if loc == "" {
			return errOutcomeUnknown
		}
		if _, err := waitOperation(env.Ctx, env, loc, purgeDeadline); err != nil {
			// The purge was submitted: whatever went wrong while watching it,
			// running it again could start a second purge.
			return fmt.Errorf("%w (%v)", errOutcomeUnknown, err)
		}
		if rounds == 1 {
			return nil
		}
		select {
		case <-env.Ctx.Done():
			return errOutcomeUnknown
		case <-time.After(settle):
		}
		est, err := runEstimate(env, path)
		if err != nil {
			return err
		}
		if est.IndexedItemCount == 0 {
			return nil
		}
		// Items on hold stay searchable after a hard delete: when the count
		// stops falling, more runs cannot help.
		if last >= 0 && est.IndexedItemCount >= last {
			return fmt.Errorf("%d message(s) still match and the count stopped falling — they are likely under a hold or retention policy", est.IndexedItemCount)
		}
		last = est.IndexedItemCount
	}
	return fmt.Errorf("messages still match after %d purge runs — check the case in Microsoft Purview", rounds)
}
