package services

import (
	"context"
	"errors"
	"fmt"
	"net/url"
	"strings"

	"swissknife-app/internal/graphapi"
	"swissknife-app/internal/ops"
	"swissknife-app/internal/session"
)

// MailboxTransferService copies a mailbox's content into a subfolder of another
// user's mailbox through the Microsoft Graph mailbox import/export APIs (v1.0,
// GA): the source folder tree is listed and recreated under a root folder in
// the target, every item is exported as an opaque full-fidelity FastTransfer
// stream and imported into the matching target folder. It is the mail
// counterpart of the OneDrive copy in the offboarding playbook.
//
// Microsoft states these APIs "aren't designed for mailbox backup and restore"
// and points at Microsoft 365 Backup for that. This is not a backup product:
// mailbox-to-mailbox migration ("copy items from one mailbox to another") is
// the documented purpose of the API, and that is exactly what this does.
//
// Endpoints (https://learn.microsoft.com/graph/api/resources/mailbox-import-export-api-overview):
//
//	GET  /users/{upn}/settings/exchange                                → primaryMailboxId (MBX:…)
//	GET  /admin/exchange/mailboxes/{mbx}/folders[/{id}/childFolders]   → mailboxFolder tree
//	GET  /admin/exchange/mailboxes/{mbx}/folders/{id}/items            → mailboxItem id + size
//	POST /admin/exchange/mailboxes/{mbx}/folders[/{id}/childFolders]   → create a folder
//	POST /admin/exchange/mailboxes/{mbx}/exportItems {itemIds: ≤ 20}   → base64 FTS per item
//	POST /admin/exchange/mailboxes/{mbx}/createImportSession           → pre-authorized importUrl
//	POST {importUrl} {FolderId, Mode: "create", Data}                  → imported item (no Bearer)
//
// Permissions are application-only in app mode: MailboxFolder.ReadWrite.All,
// MailboxItem.Read.All, MailboxItem.ImportExport.All (+ User.Read.All for the
// mailbox id lookup). Exchange Online exposes each of them as an "Application
// …" management role, so RBAC for Applications can scope the app to the
// leaver/manager mailboxes instead of the whole tenant.
type MailboxTransferService struct {
	s *session.Session
}

func NewMailboxTransferService(s *session.Session) *MailboxTransferService {
	return &MailboxTransferService{s: s}
}

// Cancel cancels the live mailbox copy: the context aborts in-flight HTTP and
// Copy returns its partial result with Canceled = true.
func (m *MailboxTransferService) Cancel() { m.s.Ops.CancelKind(ops.KindMailbox) }

const (
	// exportBatchItems is the documented maximum of itemIds per exportItems call.
	exportBatchItems = 20
	// exportBatchBytes caps the summed item sizes per exportItems call. The
	// response carries every stream base64-inline (~1.33× the item size) and
	// is read fully into memory, and Graph must answer within the client's
	// 30 s response-header timeout — 20 mails is fine, 20 huge attachments
	// is not. A single item bigger than the cap still travels alone.
	exportBatchBytes = 24 << 20
	// itemsPageSize is the $top for item listings (ids and sizes only).
	itemsPageSize = 200
)

// MailboxFolderInfo is one folder of the source mailbox as the preview sees it.
type MailboxFolderInfo struct {
	ID            string `json:"id"`
	Name          string `json:"name"`
	Parent        string `json:"parent"` // Path of the parent folder, "" at the top
	Path          string `json:"path"`   // display path, e.g. "Inbox/Projects"
	Type          string `json:"type"`   // folder class, e.g. IPF.Note
	WellKnownName string `json:"wellKnownName,omitempty"`
	Kind          string `json:"kind"` // mail | contacts | calendar | other
	Items         int    `json:"items"`
	System        bool   `json:"system"` // never copied (junk, deleted items, search folders, …)
}

// MailboxPreview is the scope of a copy: the folder tree with item counts,
// grouped by what the copy options can include. The folder listing carries
// no byte sizes, so the preview reports counts only.
type MailboxPreview struct {
	MailboxID       string              `json:"mailboxId"`
	DisplayName     string              `json:"displayName"`
	Folders         []MailboxFolderInfo `json:"folders"`
	MailFolders     int                 `json:"mailFolders"`
	MailItems       int                 `json:"mailItems"`
	ContactFolders  int                 `json:"contactFolders"`
	ContactItems    int                 `json:"contactItems"`
	CalendarFolders int                 `json:"calendarFolders"`
	CalendarItems   int                 `json:"calendarItems"`
	OtherFolders    int                 `json:"otherFolders"` // tasks, notes, journal — never copied
	OtherItems      int                 `json:"otherItems"`
	SystemFolders   int                 `json:"systemFolders"`
}

// MailboxCopyRequest describes one mailbox copy.
type MailboxCopyRequest struct {
	Source string `json:"source"`
	Target string `json:"target"`
	// Folder is the root subfolder created in the target; empty means
	// "Archive - <source display name or UPN>".
	Folder          string `json:"folder"`
	IncludeContacts bool   `json:"includeContacts"`
	IncludeCalendar bool   `json:"includeCalendar"`
	// Confirm is accepted for form symmetry with the destructive actions but
	// not checked: the copy is additive, so GuardWrite is the only gate.
	Confirm string `json:"confirm"`
}

// MailboxCopyResult is the outcome of a mailbox copy.
type MailboxCopyResult struct {
	RootFolder string            `json:"rootFolder"`
	Folders    int               `json:"folders"`    // source folders recreated in the target
	TotalItems int               `json:"totalItems"` // items in scope (from the folder counts)
	Copied     int               `json:"copied"`
	Failed     map[string]string `json:"failed"` // "Folder/Path · itemId" → reason
	Canceled   bool              `json:"canceled"`
}

// mbxFolder is the mailboxFolder resource as listed.
type mbxFolder struct {
	ID               string `json:"id"`
	DisplayName      string `json:"displayName"`
	ChildFolderCount int    `json:"childFolderCount"`
	TotalItemCount   int    `json:"totalItemCount"`
	WellKnownName    string `json:"wellKnownName"`
	Type             string `json:"type"`
}

// mbxItem is the slice of mailboxItem the copy needs.
type mbxItem struct {
	ID   string `json:"id"`
	Size int64  `json:"size"`
}

// exportedItem is one exportItemResponse entry.
type exportedItem struct {
	ItemID    string `json:"itemId"`
	ChangeKey string `json:"changeKey"`
	Data      string `json:"data"` // base64 FastTransfer stream — opaque, passed through as-is
	Error     *struct {
		Code    string `json:"code"`
		Message string `json:"message"`
	} `json:"error"`
}

func mailboxPath(mbx string) string { return "/admin/exchange/mailboxes/" + url.PathEscape(mbx) }

// folderKind maps a folder class to what the copy options talk about.
func folderKind(class string) string {
	switch {
	case class == "IPF.Note" || strings.HasPrefix(class, "IPF.Note."):
		return "mail"
	case class == "IPF.Contact" || strings.HasPrefix(class, "IPF.Contact."):
		return "contacts"
	case class == "IPF.Appointment" || strings.HasPrefix(class, "IPF.Appointment."):
		return "calendar"
	}
	return "other"
}

// kindClass is the base folder class a target folder of that kind is created with.
var kindClass = map[string]string{"mail": "IPF.Note", "contacts": "IPF.Contact", "calendar": "IPF.Appointment"}

// transparentFolder reports a container the listing may surface above the
// visible folders (the store root): descend into it, never recreate it.
func transparentFolder(f mbxFolder) bool {
	switch strings.ToLower(f.WellKnownName) {
	case "root", "msgfolderroot":
		return true
	}
	switch f.DisplayName {
	case "Top of Information Store", "IPM_SUBTREE":
		return true
	}
	return false
}

// systemWellKnown are well-known folders a leaver archive has no use for and
// whose contents would only duplicate (search folders) or pollute (junk,
// deleted items, sync issues) the manager's mailbox.
var systemWellKnown = map[string]bool{
	"junkemail": true, "deleteditems": true, "outbox": true, "syncissues": true, "conflicts": true,
	"localfailures": true, "serverfailures": true, "searchfolders": true,
	"recoverableitemsroot": true, "recoverableitemsdeletions": true, "recoverableitemspurges": true,
	"recoverableitemsversions": true, "recoverableitemsdiscoveryholds": true,
}

// systemNames catches the same folders when wellKnownName is null (non-IPM
// containers and search-folder roots carry none). Heuristic by design.
var systemNames = map[string]bool{
	"Finder": true, "Recoverable Items": true, "Common Views": true, "Views": true, "Shortcuts": true,
	"Schedule": true, "Reminders": true, "Spooler Queue": true, "Deferred Action": true,
	"Freebusy Data": true, "System": true, "Location": true, "ExchangeSyncData": true,
	"Sync Issues": true, "Junk Email": true, "Deleted Items": true, "Conversation Action Settings": true,
	"Quick Step Settings": true, "Yammer Root": true, "ApplicationDataRoot": true,
	"MailboxAssociations": true, "GraphStore": true, "PeopleConnect": true, "Sharing": true,
	"Audits": true, "Calendar Logging": true,
}

func systemFolder(f mbxFolder) bool {
	return systemWellKnown[strings.ToLower(f.WellKnownName)] || systemNames[f.DisplayName]
}

// includeFolder decides whether a folder travels under the request's options.
func includeFolder(f MailboxFolderInfo, req MailboxCopyRequest) bool {
	if f.System {
		return false
	}
	switch f.Kind {
	case "mail":
		return true
	case "contacts":
		return req.IncludeContacts
	case "calendar":
		return req.IncludeCalendar
	}
	return false
}

// splitExportBatches groups items into exportItems calls: at most
// exportBatchItems per call and roughly exportBatchBytes of item data.
func splitExportBatches(items []mbxItem) [][]mbxItem {
	var out [][]mbxItem
	var cur []mbxItem
	var size int64
	for _, it := range items {
		if len(cur) > 0 && (len(cur) >= exportBatchItems || size+it.Size > exportBatchBytes) {
			out = append(out, cur)
			cur, size = nil, 0
		}
		cur = append(cur, it)
		size += it.Size
	}
	if len(cur) > 0 {
		out = append(out, cur)
	}
	return out
}

// fatalCopyErr reports an error that makes continuing pointless: the operator
// cancelled, or the app lacks access (401/403) — every further item would fail
// the same way and bury the one actionable hint under thousands of lines.
func fatalCopyErr(err error) bool {
	if err == nil {
		return false
	}
	if errors.Is(err, context.Canceled) || errors.Is(err, context.DeadlineExceeded) {
		return true
	}
	var ge *graphapi.GraphError
	return errors.As(err, &ge) && (ge.StatusCode == 401 || ge.StatusCode == 403)
}

// mailboxID resolves a user's primary Exchange mailbox id (MBX:…@…), the
// identifier every import/export endpoint is keyed by.
func (m *MailboxTransferService) mailboxID(ctx context.Context, c *graphapi.Client, upn string) (string, error) {
	var es struct {
		PrimaryMailboxID string `json:"primaryMailboxId"`
	}
	if err := c.Get(ctx, "/users/"+url.PathEscape(upn)+"/settings/exchange", nil, &es); err != nil {
		return "", err
	}
	if es.PrimaryMailboxID == "" {
		return "", fmt.Errorf("%s has no Exchange Online mailbox", upn)
	}
	return es.PrimaryMailboxID, nil
}

// walkFolders lists the mailbox folder tree depth-first (parents before
// children, the order the copy recreates them in).
func (m *MailboxTransferService) walkFolders(ctx context.Context, c *graphapi.Client, mbx string) ([]MailboxFolderInfo, error) {
	var out []MailboxFolderInfo
	var walk func(listPath, parentPath string) error
	walk = func(listPath, parentPath string) error {
		folders, err := graphapi.ListAllInto[mbxFolder](ctx, c, listPath, url.Values{"$top": {"200"}}, 0)
		if err != nil {
			return err
		}
		for _, f := range folders {
			if ctx.Err() != nil {
				return ctx.Err()
			}
			children := mailboxPath(mbx) + "/folders/" + url.PathEscape(f.ID) + "/childFolders"
			if transparentFolder(f) {
				if f.ChildFolderCount > 0 {
					if err := walk(children, parentPath); err != nil {
						return err
					}
				}
				continue
			}
			path := f.DisplayName
			if parentPath != "" {
				path = parentPath + "/" + f.DisplayName
			}
			info := MailboxFolderInfo{
				ID: f.ID, Name: f.DisplayName, Parent: parentPath, Path: path, Type: f.Type,
				WellKnownName: f.WellKnownName, Kind: folderKind(f.Type), Items: f.TotalItemCount,
				System: systemFolder(f),
			}
			out = append(out, info)
			// A system folder takes its whole subtree with it (Deleted Items
			// keeps the deleted folders, Finder holds the search folders).
			if info.System || f.ChildFolderCount == 0 {
				continue
			}
			if err := walk(children, path); err != nil {
				return err
			}
		}
		return nil
	}
	if err := walk(mailboxPath(mbx)+"/folders", ""); err != nil {
		return nil, err
	}
	return out, nil
}

func (m *MailboxTransferService) preview(ctx context.Context, c *graphapi.Client, upn, mbx string) (*MailboxPreview, error) {
	folders, err := m.walkFolders(ctx, c, mbx)
	if err != nil {
		return nil, err
	}
	var who struct {
		DisplayName string `json:"displayName"`
	}
	_ = c.Get(ctx, "/users/"+url.PathEscape(upn), url.Values{"$select": {"displayName"}}, &who)
	pv := &MailboxPreview{MailboxID: mbx, DisplayName: who.DisplayName, Folders: folders}
	for _, f := range folders {
		switch {
		case f.System:
			pv.SystemFolders++
		case f.Kind == "mail":
			pv.MailFolders++
			pv.MailItems += f.Items
		case f.Kind == "contacts":
			pv.ContactFolders++
			pv.ContactItems += f.Items
		case f.Kind == "calendar":
			pv.CalendarFolders++
			pv.CalendarItems += f.Items
		default:
			pv.OtherFolders++
			pv.OtherItems += f.Items
		}
	}
	return pv, nil
}

// Preview lists the source mailbox's folder tree with item counts so the
// operator sees the scope before running. Read-only.
func (m *MailboxTransferService) Preview(sourceUpn string) (*MailboxPreview, error) {
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	ctx := m.s.Ctx()
	mbx, err := m.mailboxID(ctx, c, sourceUpn)
	if err != nil {
		return nil, wrapOpErr(err)
	}
	pv, err := m.preview(ctx, c, sourceUpn, mbx)
	return pv, wrapOpErr(err)
}

// Copy runs the mailbox copy as a top-level operation.
func (m *MailboxTransferService) Copy(req MailboxCopyRequest) (*MailboxCopyResult, error) {
	res, err := m.copyCtx(m.s.Ctx(), req, nil)
	return res, wrapOpErr(err)
}

func (m *MailboxTransferService) emitOverall(op *ops.Operation, done, total int, doneBytes int64, folder string) {
	emitOp(m.s.Ctx(), op, "transfer:overall", map[string]any{
		"doneItems": done, "totalItems": total, "folder": folder,
		"doneBytes": doneBytes, "totalBytes": int64(0),
		// legacy names so a console that only knows the OneDrive shape still moves
		"files": done, "totalFiles": total,
	})
}

func (m *MailboxTransferService) emitFolder(op *ops.Operation, name, status, reason string, copied int) {
	emitOp(m.s.Ctx(), op, "transfer:file", map[string]any{
		"name": name, "status": status, "reason": reason, "copied": copied,
	})
}

// copyCtx is Copy with an explicit parent context: the offboarding playbook
// passes its operation context so cancelling the playbook cancels the copy,
// and may pass a preview it already took.
func (m *MailboxTransferService) copyCtx(parent context.Context, req MailboxCopyRequest, prev *MailboxPreview) (res *MailboxCopyResult, err error) {
	if err := m.s.GuardWrite(); err != nil {
		return nil, err
	}
	if req.Source == "" || req.Target == "" {
		return nil, errors.New("source and target users are required")
	}
	if strings.EqualFold(strings.TrimSpace(req.Source), strings.TrimSpace(req.Target)) {
		return nil, errors.New("source and target mailboxes must differ")
	}
	c, err := m.s.Client()
	if err != nil {
		return nil, err
	}
	op, err := m.s.Ops.Start(parent, ops.KindMailbox)
	if err != nil {
		return nil, err
	}
	defer m.s.Ops.Finish(op)
	ctx := op.Ctx
	emitOp(m.s.Ctx(), op, "op:start", map[string]any{"target": req.Source + " → " + req.Target})

	srcMbx, err := m.mailboxID(ctx, c, req.Source)
	if err != nil {
		return nil, err
	}
	tgtMbx, err := m.mailboxID(ctx, c, req.Target)
	if err != nil {
		return nil, err
	}
	if srcMbx == tgtMbx {
		return nil, errors.New("source and target resolve to the same mailbox")
	}
	if prev == nil || prev.MailboxID != srcMbx {
		if prev, err = m.preview(ctx, c, req.Source, srcMbx); err != nil {
			return nil, err
		}
	}

	rootName := strings.TrimSpace(req.Folder)
	if rootName == "" {
		name := prev.DisplayName
		if name == "" {
			name = req.Source
		}
		rootName = "Archive - " + name
	}

	var plan []MailboxFolderInfo
	total := 0
	for _, f := range prev.Folders {
		if includeFolder(f, req) {
			plan = append(plan, f)
			total += f.Items
		}
	}
	res = &MailboxCopyResult{RootFolder: rootName, TotalItems: total, Failed: map[string]string{}}
	defer func() {
		if res != nil {
			m.s.Record("mailbox.copy", req.Source+" -> "+req.Target,
				"folder="+rootName+" folders="+itoa(res.Folders)+" copied="+itoa(res.Copied)+
					" failed="+itoa(len(res.Failed))+" canceled="+fmt.Sprint(res.Canceled), err)
		}
	}()

	var rootID string
	rootID, rootName, err = m.createUniqueRoot(ctx, c, tgtMbx, rootName)
	if err != nil {
		if errors.Is(err, context.Canceled) {
			res.Canceled = true
			return res, nil
		}
		return nil, err
	}
	res.RootFolder = rootName
	// Source path → target folder id. A folder whose parent was not copied
	// (a mail folder under a Tasks folder, say) attaches to the nearest
	// copied ancestor, ultimately the root.
	targetIDs := map[string]string{"": rootID}
	nearest := func(path string) string {
		for {
			if id, ok := targetIDs[path]; ok {
				return id
			}
			if i := strings.LastIndex(path, "/"); i >= 0 {
				path = path[:i]
			} else {
				return rootID
			}
		}
	}

	// The import session (and its pre-authorized upload URL) stays inside graphapi.
	imp := c.NewImportSession(mailboxPath(tgtMbx) + "/createImportSession")
	done := 0
	var doneBytes int64
	m.emitOverall(op, done, total, doneBytes, "")
	for _, f := range plan {
		if op.Canceled() {
			break
		}
		fid, ferr := m.ensureFolder(ctx, c, tgtMbx, nearest(f.Parent), f.Name, kindClass[f.Kind])
		if ferr != nil {
			if fatalCopyErr(ferr) {
				if errors.Is(ferr, context.Canceled) {
					// A cancel is not a failure: hand back what was copied so far.
					res.Canceled = true
					return res, nil
				}
				res.Canceled = op.Canceled()
				return res, ferr
			}
			res.Failed[f.Path+"/"] = ferr.Error()
			m.emitFolder(op, f.Path, "failed", ferr.Error(), res.Copied)
			done += f.Items
			m.emitOverall(op, done, total, doneBytes, f.Path)
			continue
		}
		targetIDs[f.Path] = fid
		res.Folders++
		before := len(res.Failed)
		copied, cerr := m.copyFolder(ctx, op, c, srcMbx, f, fid, imp, res, &done, &doneBytes, &total)
		if cerr != nil {
			if fatalCopyErr(cerr) {
				if errors.Is(cerr, context.Canceled) {
					res.Canceled = true
					return res, nil
				}
				res.Canceled = op.Canceled()
				return res, cerr
			}
			res.Failed[f.Path+"/"] = cerr.Error()
			m.emitFolder(op, f.Path, "failed", cerr.Error(), res.Copied)
			continue
		}
		failed := len(res.Failed) - before
		label := f.Path + " (" + itoa(copied) + " items)"
		switch {
		case failed > 0:
			m.emitFolder(op, label, "failed", itoa(failed)+" item(s) failed", res.Copied)
		default:
			m.emitFolder(op, label, "copied", "", res.Copied)
		}
	}
	res.Canceled = op.Canceled()
	return res, nil
}

// copyFolder exports the folder's items in batches and imports them into the
// target folder. Per-item failures land in res.Failed; only fatal errors
// (cancel, 401/403) come back as an error.
func (m *MailboxTransferService) copyFolder(ctx context.Context, op *ops.Operation, c *graphapi.Client, srcMbx string,
	f MailboxFolderInfo, targetID string, imp *graphapi.ImportSession, res *MailboxCopyResult, done *int, doneBytes *int64, total *int) (int, error) {
	items, err := graphapi.ListAllInto[mbxItem](ctx, c, mailboxPath(srcMbx)+"/folders/"+url.PathEscape(f.ID)+"/items",
		url.Values{"$select": {"id,size"}, "$top": {itoa(itemsPageSize)}}, 0)
	if err != nil {
		return 0, err
	}
	// The folder count is a snapshot; the listing is the truth for progress.
	*total += len(items) - f.Items
	copied := 0
	fail := func(id, reason string) {
		res.Failed[f.Path+" · "+id] = reason
	}
	for _, batch := range splitExportBatches(items) {
		if op.Canceled() {
			return copied, context.Canceled
		}
		ids := make([]string, 0, len(batch))
		var bytes int64
		for _, it := range batch {
			ids = append(ids, it.ID)
			bytes += it.Size
		}
		var exported struct {
			Value []exportedItem `json:"value"`
		}
		if err := c.Post(ctx, mailboxPath(srcMbx)+"/exportItems", map[string]any{"itemIds": ids}, &exported); err != nil {
			if fatalCopyErr(err) {
				return copied, err
			}
			for _, id := range ids {
				fail(id, "export: "+err.Error())
			}
			*done += len(ids)
			*doneBytes += bytes
			m.emitOverall(op, *done, *total, *doneBytes, f.Path)
			continue
		}
		seen := map[string]bool{}
		for _, e := range exported.Value {
			seen[e.ItemID] = true
			switch {
			case e.Error != nil:
				fail(e.ItemID, "export: "+e.Error.Code+" "+e.Error.Message)
			case e.Data == "":
				fail(e.ItemID, "export returned no data")
			default:
				if _, err := importItem(ctx, imp, targetID, e.Data); err != nil {
					if fatalCopyErr(err) {
						return copied, err
					}
					fail(e.ItemID, "import: "+err.Error())
				} else {
					copied++
					res.Copied++
				}
			}
		}
		for _, id := range ids {
			if !seen[id] {
				fail(id, "export response did not include the item")
			}
		}
		*done += len(ids)
		*doneBytes += bytes
		m.emitOverall(op, *done, *total, *doneBytes, f.Path)
	}
	return copied, nil
}

// ensureFolder creates a folder of the class under parentID ("" = the mailbox
// root) and returns its id. Under a fresh run root a name clash can only be a
// retry of this very run (a transient error after the create went through), so
// the existing folder is looked up and reused.
func (m *MailboxTransferService) ensureFolder(ctx context.Context, c *graphapi.Client, mbx, parentID, name, class string) (string, error) {
	path := mailboxPath(mbx) + "/folders"
	if parentID != "" {
		path = mailboxPath(mbx) + "/folders/" + url.PathEscape(parentID) + "/childFolders"
	}
	if class == "" {
		class = "IPF.Note"
	}
	var created struct {
		ID string `json:"id"`
	}
	err := c.Post(ctx, path, map[string]any{"displayName": name, "type": class}, &created)
	if err == nil {
		if created.ID == "" {
			return "", errors.New("graph: folder create returned no id")
		}
		return created.ID, nil
	}
	if !folderExistsErr(err) {
		return "", err
	}
	existing, lerr := graphapi.ListAllInto[mbxFolder](ctx, c, path,
		url.Values{"$filter": {"displayName eq '" + strings.ReplaceAll(name, "'", "''") + "'"}, "$top": {"50"}}, 0)
	if lerr != nil {
		return "", err
	}
	for _, f := range existing {
		if f.DisplayName == name {
			return f.ID, nil
		}
	}
	return "", err
}

// folderExistsErr reports Graph's "a folder with this name already exists"
// rejection (409, or an ErrorFolderExists-style code/message).
func folderExistsErr(err error) bool {
	var ge *graphapi.GraphError
	return errors.As(err, &ge) && (ge.StatusCode == 409 || strings.Contains(strings.ToLower(ge.Code+" "+ge.Message), "exist"))
}

// createUniqueRoot creates the run's root folder at the target mailbox's top
// level and returns its id and final name. A name that already exists gets a
// numeric suffix instead of being reused: every item is imported in create
// mode, so re-filling last run's folder would duplicate everything that had
// already made it across before the failure.
func (m *MailboxTransferService) createUniqueRoot(ctx context.Context, c *graphapi.Client, mbx, name string) (string, string, error) {
	const maxTries = 50
	for n := 1; n <= maxTries; n++ {
		candidate := name
		if n > 1 {
			candidate = fmt.Sprintf("%s (%d)", name, n)
		}
		var created struct {
			ID string `json:"id"`
		}
		err := c.Post(ctx, mailboxPath(mbx)+"/folders", map[string]any{"displayName": candidate, "type": "IPF.Note"}, &created)
		if err == nil {
			if created.ID == "" {
				return "", "", errors.New("graph: folder create returned no id")
			}
			return created.ID, candidate, nil
		}
		if !folderExistsErr(err) {
			return "", "", err
		}
	}
	return "", "", fmt.Errorf("no free folder name for %q after %d tries", name, maxTries)
}

// importItem uploads one exported stream into the target folder through the
// mailbox's import session (Mode create: a new item, never an update).
func importItem(ctx context.Context, imp *graphapi.ImportSession, folderID, data string) (string, error) {
	body := map[string]any{"FolderId": folderID, "Mode": "create", "Data": data}
	var out struct {
		ItemID string `json:"itemId"`
	}
	err := imp.Post(ctx, body, &out)
	return out.ItemID, err
}
