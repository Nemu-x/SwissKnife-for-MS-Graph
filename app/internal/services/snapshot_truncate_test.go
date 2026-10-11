package services

import (
	"net/http"
	"testing"
)

func TestDiffTreatsTruncatedSectionAsNotComparable(t *testing.T) {
	mk := func(truncated bool, ids ...string) *snapshotFile {
		objs := []map[string]any{}
		for _, id := range ids {
			objs = append(objs, map[string]any{"id": id})
		}
		return &snapshotFile{
			Meta:     SnapshotMeta{Sections: []SnapshotSection{{Name: "groups", Count: len(ids), Truncated: truncated}}},
			Sections: map[string][]map[string]any{"groups": objs},
		}
	}
	// The first N groups differ between two capped reads: no drift may follow.
	d := diffSnapshots(mk(true, "g1", "g2"), mk(true, "g2", "g3"))
	if d.Added+d.Removed+d.Changed != 0 {
		t.Fatalf("truncated section produced drift: %+v", d)
	}
	var sec SectionDiff
	for _, s := range d.Sections {
		if s.Name == "groups" {
			sec = s
		}
	}
	if !sec.Skipped || !sec.Truncated || sec.Note == "" {
		t.Errorf("groups diff = %+v, want skipped+truncated with a note", sec)
	}
	// Only one side capped is not comparable either.
	if d := diffSnapshots(mk(false, "g1"), mk(true, "g1", "g2")); d.Added != 0 {
		t.Errorf("one-sided truncation produced drift: %+v", d)
	}
	// Uncapped sections still diff.
	if d := diffSnapshots(mk(false, "g1"), mk(false, "g1", "g2")); d.Added != 1 {
		t.Errorf("added = %d, want 1", d.Added)
	}
}

func TestListSectionMarksTruncation(t *testing.T) {
	sess := harness(t, func(w http.ResponseWriter, r *http.Request) {
		w.Write([]byte(`{"value":[{"id":"a"},{"id":"b"},{"id":"c"}]}`))
	})
	c, err := sess.Client()
	if err != nil {
		t.Fatal(err)
	}
	objs, truncated, err := listSection("/groups", nil, 2)(sess.Ctx(), c)
	if err != nil || len(objs) != 2 || !truncated {
		t.Errorf("capped = %d %v %v, want 2 objects truncated", len(objs), truncated, err)
	}
	if _, truncated, _ := listSection("/groups", nil, 0)(sess.Ctx(), c); truncated {
		t.Error("uncapped read reported truncation")
	}
}
