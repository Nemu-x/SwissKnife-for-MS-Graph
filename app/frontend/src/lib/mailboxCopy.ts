import { useCallback, useRef } from 'react'
import { useStore } from './store'
import { api, errMessage } from './api'
import type { services } from '../../wailsjs/go/models'

export type MailboxCopyParams = {
  source: string
  target: string
  folder: string
  includeContacts: boolean
  includeCalendar: boolean
}

// The mailbox copy runs as the "mailbox" job in the store: the backend stamps
// its events with opKind "mailbox", so progress routes to that key and the
// console keeps updating while the operator is on another page.
export function useMailboxCopy() {
  const { jobs, patchJob, jobLog, toast, clearJob } = useStore()
  const job = jobs.mailbox
  // Generation of the request that owns the slot: React applies the `running`
  // patch asynchronously, so two clicks in a row could both pass a state check.
  // The ref is claimed synchronously and only the owner may write the outcome.
  const owner = useRef(0)
  const seq = useRef(0)

  const start = useCallback(async (p: MailboxCopyParams): Promise<services.MailboxCopyResult | null> => {
    if (owner.current !== 0 || jobs.mailbox?.running) return null
    const gen = ++seq.current
    owner.current = gen
    patchJob('mailbox', {
      running: true, canceled: false, progress: 'Starting…', log: [], result: null, error: null,
      startedAt: Date.now(), opId: undefined,
    })
    jobLog('mailbox', `▶ Mailbox ${p.source} → ${p.target}`)
    try {
      const r = await api.mailboxTransfer.copy({ ...p, confirm: '' })
      if (owner.current !== gen) return r
      patchJob('mailbox', { result: r })
      const failed = Object.keys(r.failed || {}).length
      jobLog('mailbox', `${r.canceled ? '⏹ Canceled' : '✓ Done'} — ${r.copied} copied in ${r.folders} folder(s), ${failed} failed → ${r.rootFolder}`)
      toast(r.canceled ? 'info' : 'ok', `${r.copied} copied${r.canceled ? ' (canceled)' : ''}`)
      return r
    } catch (e) {
      if (owner.current !== gen) return null
      patchJob('mailbox', { error: errMessage(e) })
      jobLog('mailbox', `✗ Error: ${errMessage(e)}`)
      return null
    } finally {
      if (owner.current === gen) {
        owner.current = 0
        patchJob('mailbox', { running: false, progress: '' })
      }
    }
  }, [jobs.mailbox?.running, patchJob, jobLog, toast])

  const cancel = useCallback(() => {
    patchJob('mailbox', { canceled: true, progress: 'Canceling…' })
    jobLog('mailbox', '⏹ Cancel requested — stopping after the current batch…')
    api.mailboxTransfer.cancel().catch(() => {})
  }, [patchJob, jobLog])

  const clear = useCallback(() => clearJob('mailbox'), [clearJob])

  return { job, start, cancel, clear }
}
