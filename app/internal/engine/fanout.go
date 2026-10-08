package engine

// Multi-tenant reads: a read action can run against a connection other than
// the session's (another saved profile). Only Graph implementations take
// part — PowerShell hosts and Exchange probes belong to the session's own
// connection.

// graphReader returns the action's Graph read implementation, if any.
func graphReader(a Action) Reader {
	if a.Danger != Read {
		return nil
	}
	for _, impl := range a.Impls {
		if r, ok := impl.(readerImpl); ok && impl.Backend() == BackendGraph {
			return r.Reader
		}
	}
	return nil
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
