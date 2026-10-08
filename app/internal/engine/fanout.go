package engine

// Multi-tenant reads: a read action can run against a connection other than
// the session's (another saved profile). Only Graph implementations take
// part — PowerShell hosts and Exchange probes belong to the session's own
// connection.

// graphReader returns the action's Graph read implementation, if any.
// Reads that take a user or group are left out: the picked object exists in
// the connected tenant only.
func graphReader(a Action) Reader {
	if a.Danger != Read {
		return nil
	}
	for _, f := range a.Fields {
		if f.Kind == FieldUser || f.Kind == FieldGroup {
			return nil
		}
	}
	for _, impl := range a.Impls {
		if r, ok := impl.(readerImpl); ok && impl.Backend() == BackendGraph {
			return r.Reader
		}
	}
	return nil
}

// CheckFanOut validates a cross-tenant run once, before any tenant is read.
func (e *Engine) CheckFanOut(actionID string, in Inputs) error {
	a, err := e.lookup(actionID)
	if err != nil {
		return err
	}
	if graphReader(a) == nil {
		return &Error{Code: "notFanOut", Msg: "this action cannot run across tenants"}
	}
	return validate(a.Manifest, cloneInputs(in))
}

// RunOn runs a read action on env (another tenant's connection).
func (e *Engine) RunOn(env Env, actionID string, in Inputs) (*ReadResult, error) {
	a, err := e.lookup(actionID)
	if err != nil {
		return nil, err
	}
	r := graphReader(a)
	if r == nil {
		return nil, &Error{Code: "notFanOut", Msg: "this action cannot run across tenants"}
	}
	in = cloneInputs(in)
	if err := validate(a.Manifest, in); err != nil {
		return nil, err
	}
	res, err := r.Read(env, in)
	if err != nil {
		return nil, e.WrapErr(err)
	}
	if res == nil {
		res = &ReadResult{}
	}
	res.Backend = BackendGraph
	return res, nil
}
