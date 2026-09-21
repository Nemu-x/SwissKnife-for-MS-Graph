import { useCallback } from 'react'
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

  const start = useCallback(async (p: MailboxCopyParams): Promise<services.MailboxCopyResult | null> => {
    patchJob('mailbox', {
      running: true, canceled: false, progress: 'Starting…', log: [], result: null, error: null,
      startedAt: Date.now(), opId: undefined,
    })
    jobLog('mailbox', `▶ Mailbox ${p.source} → ${p.target}`)
    try {
      const r = await api.mailboxTransfer.copy({ ...p, confirm: '' })
      patchJob('mailbox', { result: r })
      const failed = Object.keys(r.failed || {}).length
      jobLog('mailbox', `${r.canceled ? '⏹ Canceled' : '✓ Done'} — ${r.copied} copied in ${r.folders} folder(s), ${failed} failed → ${r.rootFolder}`)
      toast(r.canceled ? 'info' : 'ok', `${r.copied} copied${r.canceled ? ' (canceled)' : ''}`)
      return r
    } catch (e) {
      patchJob('mailbox', { error: errMessage(e) })
      jobLog('mailbox', `✗ Error: ${errMessage(e)}`)
      return null
    } finally {
      patchJob('mailbox', { running: false, progress: '' })
    }
  }, [patchJob, jobLog, toast])

  const cancel = useCallback(() => {
    patchJob('mailbox', { canceled: true, progress: 'Canceling…' })
    jobLog('mailbox', '⏹ Cancel requested — stopping after the current batch…')
    api.mailboxTransfer.cancel().catch(() => {})
  }, [patchJob, jobLog])

  const clear = useCallback(() => clearJob('mailbox'), [clearJob])

  return { job, start, cancel, clear }
}
