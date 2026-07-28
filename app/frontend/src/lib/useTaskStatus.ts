import { useCallback, useRef, useState } from 'react'
import { useStore } from './store'
import { errMessage } from './api'
import type { ActionStatus } from '../components/TaskPage'

// Every task page runs writes the same way: perform the call, record the outcome
// on the tile that started it (a toast disappears, the tile does not), and toast
// it too. Having one implementation keeps the semantics identical everywhere —
// the copies had already started to drift on the error path.
export interface TaskStatus {
  status: Record<string, ActionStatus>
  /**
   * True while a write is in flight. Writes go straight to the API rather than
   * through useAsync, so without this the page header showed no busy state and
   * the button stayed clickable for the whole call.
   */
  busy: boolean
  /** Record an outcome without running anything (for callers with their own flow). */
  mark: (id: string, ok: boolean, text: string) => void
  /** Run a write; on success record `okText`, on failure record the error. */
  doWrite: (id: string, fn: () => Promise<unknown>, okText: string) => Promise<void>
  /** Like doWrite, but the call's result is handed back for display. */
  doShow: <T>(id: string, fn: () => Promise<T>, okText: string) => Promise<T | undefined>
}

export function useTaskStatus(): TaskStatus {
  const { toast } = useStore()
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})
  const [inFlight, setInFlight] = useState(0)

  const mark = useCallback((id: string, ok: boolean, text: string) => {
    setStatus((s) => ({ ...s, [id]: { ok, text, at: Date.now() } }))
  }, [])

  // One tile can drive several calls (Add and Remove share an id), so a slow
  // first request must not land on top of a faster second one and report the
  // wrong operation. Only the newest call per id may write the status.
  const seq = useRef<Record<string, number>>({})

  const doShow = useCallback(async <T,>(id: string, fn: () => Promise<T>, okText: string) => {
    const ticket = (seq.current[id] ?? 0) + 1
    seq.current[id] = ticket
    const current = () => seq.current[id] === ticket
    setInFlight((n) => n + 1)
    try {
      const r = await fn()
      if (current()) {
        mark(id, true, okText)
        toast('ok', okText)
      }
      return r
    } catch (e) {
      const m = errMessage(e)
      if (current()) {
        mark(id, false, m)
        toast('err', m)
      }
      return undefined
    } finally {
      setInFlight((n) => n - 1)
    }
  }, [mark, toast])

  const doWrite = useCallback(async (id: string, fn: () => Promise<unknown>, okText: string) => {
    await doShow(id, fn, okText)
  }, [doShow])

  return { status, busy: inFlight > 0, mark, doWrite, doShow }
}
