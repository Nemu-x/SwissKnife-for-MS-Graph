package engine

import "context"

// Row is one line of a read action's answer: column → display value.
type Row map[string]string

// ReadResult is a read action's answer: rows plus the column order to show
// them in (map keys have none). Column names are i18n tokens on the frontend
// (actions.columns.<name>).
type ReadResult struct {
	Columns []string `json:"columns"`
	Rows    []Row    `json:"rows"`
	Backend Backend  `json:"backend"`
	// Note is an optional line under the table (an i18n key + params), e.g.
	// how many mailboxes a scan covered.
	Note *Reason `json:"note,omitempty"`
	// TenantNotes are the notes of each tenant in a cross-tenant run.
	TenantNotes []TenantNote `json:"tenantNotes,omitempty"`
}

// TenantNote is one tenant's note in a cross-tenant run.
type TenantNote struct {
	Tenant string `json:"tenant"`
	Note   Reason `json:"note"`
}

// Reader is implemented by read actions (Danger Read). They answer with rows
// instead of a plan: nothing changes, so there is nothing to preview.
type Reader interface {
	Backend() Backend
	Read(env Env, in Inputs) (*ReadResult, error)
}

// readerImpl lets a Reader sit in Action.Impls next to write implementations.
type readerImpl struct{ Reader }

func (readerImpl) Plan(Env, Inputs) ([]Change, error) { return nil, errReadOnlyAction }
func (readerImpl) Apply(Env, Inputs, Change) error      { return errReadOnlyAction }

var errReadOnlyAction = &Error{Code: "readAction", Msg: "this action only reads — run it instead of previewing"}

// ReadImpl wraps a Reader for Action.Impls.
func ReadImpl(r Reader) Impl { return readerImpl{r} }

// Run executes a read action. Reads change nothing, so they need no guard,
// operation or journal entry.
func (e *Engine) Run(ctx context.Context, actionID string, in Inputs) (*ReadResult, error) {
	a, err := e.lookup(actionID)
	if err != nil {
		return nil, err
	}
	if a.Danger != Read {
		return nil, &Error{Code: "writeAction", Msg: "this action changes the tenant — preview it first"}
	}
	in = cloneInputs(in)
	if err := validate(a.Manifest, in); err != nil {
		return nil, err
	}
	impl, reason := e.resolve(a)
	if impl == nil {
		return nil, &Error{Code: "unavailable", Msg: "action unavailable: " + reason.Key}
	}
	r, ok := impl.(readerImpl)
	if !ok {
		return nil, errReadOnlyAction
	}
	res, err := r.Read(e.env(ctx), in)
	if err != nil {
		return nil, e.WrapErr(err)
	}
	if res == nil {
		res = &ReadResult{}
	}
	res.Backend = impl.Backend()
	return res, nil
}
