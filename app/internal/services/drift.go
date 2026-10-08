package services

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// Drift watch: while the app runs, the tenant's configuration is collected
// on a schedule and compared with a baseline snapshot; new differences raise
// an alert in the app and, if wanted, in a Teams channel. Nothing is written
// to the snapshot list by a check — the comparison happens in memory.

// DriftWatch is persisted as drift.json in the config directory.
type DriftWatch struct {
	BaselineID    string        `json:"baselineId"`
	EveryHours    int           `json:"everyHours"` // 0 = off
	NotifyTeams   bool          `json:"notifyTeams"`
	LastCheck     time.Time     `json:"lastCheck"`
	LastSignature string        `json:"lastSignature,omitempty"`
	// LastSections keeps each section's signature, so a section that could
	// not be read in a check keeps its previous state instead of flapping.
	LastSections map[string]string `json:"lastSections,omitempty"`
	LastSummary   *DriftSummary `json:"lastSummary,omitempty"`
	LastError     string        `json:"lastError,omitempty"`
}

// DriftSummary is what a check found.
type DriftSummary struct {
	Added    int      `json:"added"`
	Removed  int      `json:"removed"`
	Changed  int      `json:"changed"`
	Sections []string `json:"sections"`
}

var driftMu sync.Mutex // one check at a time, config writes serialized

func driftPath(dir string) string { return filepath.Join(dir, "drift.json") }

func loadDrift(dir string) DriftWatch {
	var w DriftWatch
	if b, err := os.ReadFile(driftPath(dir)); err == nil {
		_ = json.Unmarshal(b, &w)
	}
	return w
}

func saveDrift(dir string, w DriftWatch) error {
	b, err := json.MarshalIndent(w, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(driftPath(dir), b, 0o600)
}

// GetDriftWatch returns the watch settings and the last result.
func (x *SnapshotService) GetDriftWatch() (*DriftWatch, error) {
	dir := x.s.ConfigDir()
	if dir == "" {
		return nil, errors.New("config directory is not set")
	}
	driftMu.Lock()
	defer driftMu.Unlock()
	w := loadDrift(dir)
	return &w, nil
}

// SetDriftWatch changes the baseline, the interval (0, 1, 6 or 24 hours) and
// the Teams alert; the last result is kept unless the baseline changed.
func (x *SnapshotService) SetDriftWatch(cfg DriftWatch) (*DriftWatch, error) {
	dir := x.s.ConfigDir()
	if dir == "" {
		return nil, errors.New("config directory is not set")
	}
	switch cfg.EveryHours {
	case 0, 1, 6, 24:
	default:
		return nil, errors.New("check every 1, 6 or 24 hours, or turn the watch off")
	}
	if cfg.EveryHours > 0 {
		if _, err := x.loadMeta(cfg.BaselineID); err != nil {
			return nil, fmt.Errorf("baseline snapshot: %w", err)
		}
	}
	driftMu.Lock()
	defer driftMu.Unlock()
	w := loadDrift(dir)
	if w.BaselineID != cfg.BaselineID {
		w.LastSignature, w.LastSummary, w.LastCheck, w.LastSections, w.LastError = "", nil, time.Time{}, nil, ""
	}
	w.BaselineID, w.EveryHours, w.NotifyTeams = cfg.BaselineID, cfg.EveryHours, cfg.NotifyTeams
	if err := saveDrift(dir, w); err != nil {
		return nil, err
	}
	x.s.Record("drift.watch", w.BaselineID, fmt.Sprintf("everyHours=%d notifyTeams=%v", w.EveryHours, w.NotifyTeams), nil)
	return &w, nil
}

// CheckDriftNow runs a check immediately (and alerts like a scheduled one).
func (x *SnapshotService) CheckDriftNow() (*DriftSummary, error) {
	return x.checkDrift(true)
}

// sectionSignatures identifies each section's differences ("" = none), so
// the same drift alerts once.
func sectionSignatures(d *SnapshotDiff) map[string]string {
	out := map[string]string{}
	for _, sec := range d.Sections {
		var keys []string
		for _, kind := range []struct {
			tag  string
			list []ObjectChange
		}{{"+", sec.Added}, {"-", sec.Removed}, {"~", sec.Changed}} {
			for _, o := range kind.list {
				line := sec.Name + kind.tag + o.Key
				if kind.tag == "~" {
					b, _ := json.Marshal(o.Changes)
					line += string(b)
				}
					keys = append(keys, line)
			}
		}
		if len(keys) > 0 {
			sort.Strings(keys)
			sum := sha256.Sum256([]byte(strings.Join(keys, "\n")))
			out[sec.Name] = hex.EncodeToString(sum[:])
		}
	}
	return out
}

// combineSignatures folds per-section signatures into one ("" = no drift).
func combineSignatures(m map[string]string) string {
	var keys []string
	for k, v := range m {
		if v != "" {
			keys = append(keys, k+"="+v)
		}
	}
	if len(keys) == 0 {
		return ""
	}
	sort.Strings(keys)
	sum := sha256.Sum256([]byte(strings.Join(keys, "\n")))
	return hex.EncodeToString(sum[:])
}

func (x *SnapshotService) checkDrift(manual bool) (*DriftSummary, error) {
	dir := x.s.ConfigDir()
	c, err := x.s.Client()
	if err != nil {
		return nil, err
	}
	driftMu.Lock()
	defer driftMu.Unlock()
	w := loadDrift(dir)
	if w.BaselineID == "" {
		return nil, errors.New("choose a baseline snapshot first")
	}
	fail := func(err error) (*DriftSummary, error) {
		w.LastCheck, w.LastError = time.Now(), err.Error()
		_ = saveDrift(dir, w)
		return nil, err
	}
	baseline, err := x.load(w.BaselineID)
	if err != nil {
		return fail(fmt.Errorf("baseline snapshot: %w", err))
	}
	if err := sameTenant(baseline.Meta, x.s.TenantID()); err != nil {
		return fail(fmt.Errorf("baseline: %w", err))
	}
	op, err := x.s.Ops.Start(x.s.Ctx(), ops.KindDrift)
	if err != nil {
		return nil, err // a check is running; the next tick tries again
	}
	defer x.s.Ops.Finish(op)

	now, _, err := x.collect(op.Ctx, c, nil)
	w.LastCheck = time.Now()
	if err != nil {
		w.LastError = wrapOpErr(err).Error()
		_ = saveDrift(dir, w)
		return nil, err
	}
	now.Meta.Name, now.Meta.TakenAt = "now", w.LastCheck
	diff := diffSnapshots(baseline, now)
	sum := &DriftSummary{Added: diff.Added, Removed: diff.Removed, Changed: diff.Changed, Sections: []string{}}
	for _, s := range diff.Sections {
		if len(s.Added)+len(s.Removed)+len(s.Changed) > 0 {
			sum.Sections = append(sum.Sections, s.Name)
		}
	}
	secSigs := sectionSignatures(diff)
	for _, info := range now.Meta.Sections {
		if info.Skipped {
			if prev, ok := w.LastSections[info.Name]; ok {
				secSigs[info.Name] = prev
			}
		}
	}
	sig := combineSignatures(secSigs)
	fresh := sig != "" && sig != w.LastSignature
	w.LastSignature, w.LastSections, w.LastSummary, w.LastError = sig, secSigs, sum, ""
	if err := saveDrift(dir, w); err != nil {
		return nil, err
	}
	if fresh || (manual && sig != "") {
		emitEvent(x.s.Ctx(), "drift:detected", map[string]any{
			"added": sum.Added, "removed": sum.Removed, "changed": sum.Changed, "sections": sum.Sections,
		})
	}
	if fresh && w.NotifyTeams {
		if cfg := loadNotifyConfig(dir); cfg != nil && cfg.WebhookURL != "" {
			facts := [][2]string{
				{"Baseline", baseline.Meta.Name + " (" + baseline.Meta.TakenAt.Format("2006-01-02") + ")"},
				{"Tenant", x.s.ProfileName()},
				{"Added / removed / changed", fmt.Sprintf("%d / %d / %d", sum.Added, sum.Removed, sum.Changed)},
				{"Sections", strings.Join(sum.Sections, ", ")},
			}
			_ = postAdaptiveCard(x.s.Ctx(), cfg.WebhookURL, "Configuration drift detected", facts, nil)
		}
	}
	x.s.Record("drift.check", w.BaselineID, fmt.Sprintf("added=%d removed=%d changed=%d", sum.Added, sum.Removed, sum.Changed), nil)
	return sum, nil
}

// StartDriftWatcher runs scheduled checks until ctx ends. It wakes every few
// minutes and checks when the interval has passed and a tenant is connected.
func StartDriftWatcher(ctx context.Context, s *session.Session) {
	x := NewSnapshotService(s)
	go func() {
		t := time.NewTicker(5 * time.Minute)
		defer t.Stop()
		for {
			select {
			case <-ctx.Done():
				return
			case <-t.C:
			}
			dir := s.ConfigDir()
			if dir == "" || !s.Connected() {
				continue
			}
			driftMu.Lock()
			w := loadDrift(dir)
			driftMu.Unlock()
			if w.EveryHours == 0 || w.BaselineID == "" || time.Since(w.LastCheck) < time.Duration(w.EveryHours)*time.Hour {
				continue
			}
			_, _ = x.checkDrift(false)
		}
	}()
}
