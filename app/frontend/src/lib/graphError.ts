// Parsing of backend error strings into structure. Two formats arrive across
// the Wails boundary:
//   1. "operr:{json}"  — the services.OpError envelope (code/status/requestId/hint)
//   2. "graph: <status> <code>: <message> (requestId=<id>)" — a raw GraphError
// Anything else is passed through as a plain message.

export type ParsedError = {
  message: string
  code?: string
  status?: number
  requestId?: string
  hint?: string // missing Graph application permission, when known
}

const GRAPH_RE = /^graph: (\d{3}) (\S+): ([\s\S]*?)(?: \(requestId=([^)]+)\))?$/

export function parseErr(raw: string): ParsedError {
  if (raw.startsWith('operr:')) {
    try {
      const o = JSON.parse(raw.slice(6))
      return {
        message: o.message || raw,
        code: o.code || undefined,
        status: o.status || undefined,
        requestId: o.requestId || undefined,
        hint: o.hint || undefined,
      }
    } catch {
      /* fall through to plain */
    }
  }
  const m = GRAPH_RE.exec(raw)
  if (m) {
    return { message: m[3], code: m[2], status: Number(m[1]), requestId: m[4] || undefined }
  }
  return { message: raw }
}
