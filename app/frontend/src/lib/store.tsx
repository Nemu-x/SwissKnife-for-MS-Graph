import { createContext, useContext, useEffect, useState, type ReactNode, useCallback, useRef } from 'react'
import { api, errMessage, type Status } from './api'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import i18n from '../i18n'
import { applyAccent } from './color'
import { humanBytes } from './format'

type Theme = 'dark' | 'light'
type Toast = { id: number; kind: 'ok' | 'err' | 'info'; text: string }

// A long-running background operation (offboarding copy, cleanup scan). It lives
// in the store — not in a page — so it survives navigating away and back, and
// its live progress/log keep updating even while the page is unmounted.
export type JobState = {
  running: boolean
  canceled: boolean
  progress: string // one-line live status
  log: string[]
  result: any // page-specific payload
  error: string | null
  startedAt: number
  steps?: PlaybookLiveStep[] // playbook job: live per-step status
  opId?: string // backend operation id (stamped by the op:start event)
}
// One playbook step as streamed from the backend ("playbook:step" events).
export type PlaybookLiveStep = {
  name: string
  detail?: string
  running?: boolean
  ok?: boolean
  error?: string
  errorCode?: string
  hint?: string // missing Graph permission (403), when the backend derived one
}
export type TransferParams = { source: string; target: string; dest: string; overwrite: boolean; usePool: boolean; pool: string }
export type CleanupParams = { mode: 'duplicates' | 'versions'; ownerType: 'user' | 'site'; ownerId: string }
// One row of a bulk CSV run: a human label plus the API call to make.
export type BulkItem = { label: string; run: () => Promise<string | void> }
export type BulkRowResult = { label: string; ok: boolean; detail?: string; error?: string }

const emptyJob = (): JobState => ({ running: false, canceled: false, progress: '', log: [], result: null, error: null, startedAt: 0 })
const stamp = (s: string) => `${new Date().toLocaleTimeString()}  ${s}`

interface Store {
  status: Status | null
  setStatus: (s: Status | null) => void
  refreshStatus: () => Promise<void>
  connected: boolean
  readOnly: boolean

  domains: string[]
  loadDomains: () => Promise<void>

  theme: Theme
  toggleTheme: () => void

  accent: string
  setAccent: (hex: string) => void

  safeMode: boolean
  setSafeMode: (v: boolean) => void

  access: Record<string, boolean>
  hideUnavailable: boolean
  setHideUnavailable: (v: boolean) => void
  checkAccess: () => Promise<void>

  jobs: Record<string, JobState>
  jobLog: (key: string, line: string) => void
  patchJob: (key: string, patch: Partial<JobState>) => void
  startTransfer: (p: TransferParams) => Promise<void>
  cancelTransfer: () => void
  startCleanupScan: (p: CleanupParams) => Promise<void>
  cancelCleanupScan: () => void
  startBulkRun: (items: BulkItem[]) => Promise<void>
  cancelBulkRun: () => void
  startPlaybook: (kind: 'onboard' | 'offboard', target: string, call: () => Promise<any>) => Promise<any>
  cancelPlaybook: () => void
  clearJob: (key: string) => void

  // Cross-navigation result cache: pages stash expensive or one-time results
  // (usage reports, tenant scans, freshly issued secrets) so leaving the page
  // does not destroy them. Cleared on disconnect.
  cache: Record<string, any>
  setCache: (key: string, value: any) => void

  toasts: Toast[]
  toast: (kind: Toast['kind'], text: string) => void
  dismiss: (id: number) => void
}

const Ctx = createContext<Store | null>(null)

export function StoreProvider({ children }: { children: ReactNode }) {
  const [status, setStatus] = useState<Status | null>(null)
  const [theme, setTheme] = useState<Theme>((localStorage.getItem('theme') as Theme) || 'dark')
  const [safeMode, setSafeModeState] = useState<boolean>(localStorage.getItem('safeMode') !== 'false')
  const [toasts, setToasts] = useState<Toast[]>([])
  const [domains, setDomains] = useState<string[]>([])
  const [access, setAccess] = useState<Record<string, boolean>>(() => {
    try { return JSON.parse(localStorage.getItem('access') || '{}') } catch { return {} }
  })
  const [hideUnavailable, setHideUnavailableState] = useState<boolean>(localStorage.getItem('hideUnavailable') === 'true')
  // '' = Auto (use the theme palette accent); a hex overrides it.
  const [accent, setAccentState] = useState<string>(localStorage.getItem('accent') ?? '')

  const [jobs, setJobs] = useState<Record<string, JobState>>({})
  const patchJob = useCallback((key: string, patch: Partial<JobState>) => {
    setJobs((all) => ({ ...all, [key]: { ...(all[key] ?? emptyJob()), ...patch } }))
  }, [])
  const jobLog = useCallback((key: string, s: string) => {
    setJobs((all) => {
      const cur = all[key] ?? emptyJob()
      return { ...all, [key]: { ...cur, log: [...cur.log, stamp(s)] } }
    })
  }, [])
  const clearJob = useCallback((key: string) => {
    setJobs((all) => ({ ...all, [key]: emptyJob() }))
  }, [])

  const [cache, setCacheState] = useState<Record<string, any>>({})
  const setCache = useCallback((key: string, value: any) => {
    setCacheState((c) => ({ ...c, [key]: value }))
  }, [])

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme)
    localStorage.setItem('theme', theme)
  }, [theme])

  // Re-apply accent whenever it changes or the theme flips (theme resets vars).
  useEffect(() => {
    applyAccent(accent)
    localStorage.setItem('accent', accent)
  }, [accent, theme])

  const toast = useCallback((kind: Toast['kind'], text: string) => {
    const id = Date.now() + Math.random()
    setToasts((t) => [...t, { id, kind, text }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 5000)
  }, [])

  const refreshStatus = useCallback(async () => {
    try {
      setStatus(await api.connect.status())
    } catch {
      /* ignore */
    }
  }, [])

  useEffect(() => {
    refreshStatus()
  }, [refreshStatus])

  const loadDomains = useCallback(async () => {
    try {
      setDomains(await api.connect.domains())
    } catch {
      setDomains([])
    }
  }, [])

  // Clear cached domains and page results on disconnect; domain loading is
  // opt-in (checkbox on Connect).
  useEffect(() => {
    if (!status?.connected) {
      setDomains([])
      setCacheState({})
    }
  }, [status?.connected])

  // Stream backend progress into the jobs, mounted once at the app root so it
  // keeps flowing regardless of which page (if any) is currently shown.
  useEffect(() => {
    // Demultiplex events by backend operation id: the first live op of a kind
    // keeps the kind's job key (today's pages), a concurrent second op of the
    // same kind is routed to "kind:opId" so progress lines never mix.
    const opKeys = new Map<string, string>()
    const keyOf = (d: any, kind: string) => (d?.opId && opKeys.get(d.opId)) || kind
    const offOS = EventsOn('op:start', (d: any) => {
      if (!d?.opId || !d?.opKind) return
      setJobs((all) => {
        const primary = all[d.opKind]
        const taken = primary?.running && primary.opId && primary.opId !== d.opId
        const key = taken ? `${d.opKind}:${d.opId}` : d.opKind
        opKeys.set(d.opId, key)
        // Stamp the opId onto the job the operator just started; a child op
        // (e.g. a playbook's backup copy) only claims routing, not job state.
        if (!taken && primary?.running && !primary.opId) {
          return { ...all, [key]: { ...primary, opId: d.opId } }
        }
        return all
      })
    })
    const offP = EventsOn('transfer:progress', (d: any) => {
      const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
      patchJob(keyOf(d, 'transfer'), { progress: `${d.name} — ${pct}%` })
    })
    const offF = EventsOn('transfer:file', (d: any) => {
      const icon = d.status === 'copied' ? '✓' : d.status === 'skipped' ? '↷' : '✗'
      const key = keyOf(d, 'transfer')
      jobLog(key, `${icon} ${d.name}${d.reason ? ' — ' + d.reason : ''}`)
      patchJob(key, { progress: `${d.copied} copied…` })
    })
    const offC = EventsOn('cleanup:progress', (d: any) => {
      patchJob(keyOf(d, 'cleanup'), { progress: d.total > 0 ? `${d.stage} ${d.done}/${d.total}` : `${d.stage} ${d.done}…` })
    })
    const offCL = EventsOn('cleanup:log', (d: any) => {
      // Payload moved from a bare string to {text, opId} with the op envelope.
      jobLog(keyOf(d, 'cleanup'), typeof d === 'string' ? d : d.text)
    })
    // Whole-transfer percentage (bytes done vs. scanned total). Feeds both the
    // OneDrive transfer console and, when a playbook backup is running, the
    // playbook job's live progress line.
    const offO = EventsOn('transfer:overall', (d: any) => {
      const line = d.totalBytes > 0
        ? `${humanBytes(d.doneBytes)} / ${humanBytes(d.totalBytes)} — ${Math.round((d.doneBytes / d.totalBytes) * 100)}% (${d.files}/${d.totalFiles})`
        : `${d.files} file(s) processed…`
      const key = keyOf(d, 'transfer')
      setJobs((all) => {
        const next: Record<string, JobState> = { ...all, [key]: { ...(all[key] ?? emptyJob()), progress: line } }
        if (all.playbook?.running) next.playbook = { ...all.playbook, progress: line }
        return next
      })
    })
    // Live playbook steps: "running" appends a pending row, "done" resolves it.
    const offPB = EventsOn('playbook:step', (d: any) => {
      const key = keyOf(d, 'playbook')
      setJobs((all) => {
        const cur = all[key] ?? emptyJob()
        const steps = [...(cur.steps || [])]
        if (d.status === 'running') {
          steps.push({ name: d.name, detail: d.detail, running: true })
        } else {
          let i = steps.length - 1
          while (i >= 0 && !steps[i].running) i--
          const done = { name: d.name, detail: d.detail, ok: !!d.ok, error: d.error, errorCode: d.errorCode, hint: d.hint }
          if (i >= 0) steps[i] = done
          else steps.push(done)
        }
        return { ...all, [key]: { ...cur, steps } }
      })
      if (d.status === 'done') {
        jobLog(key, `${d.ok ? '✓' : '✗'} ${d.name}${d.detail ? ' — ' + d.detail : ''}${d.error ? ' — ' + d.error : ''}`)
      }
    })
    return () => { offOS(); offP(); offF(); offC(); offCL(); offO(); offPB() }
  }, [patchJob, jobLog])

  const startTransfer = useCallback(async (p: TransferParams) => {
    patchJob('transfer', { ...emptyJob(), running: true, startedAt: Date.now(), progress: 'Starting…' })
    jobLog('transfer', `▶ Copy ${p.source} → ${p.usePool ? '(pool)' : p.target}`)
    try {
      const folder = p.dest.trim() || p.source.split('@')[0]
      let tgt = p.target
      if (p.usePool) {
        const prev = await api.drive.offboardingPreview(p.source)
        const list = p.pool.split(/[\n,]/).map((s) => s.trim()).filter(Boolean)
        tgt = await api.drive.pickTarget(list, prev.totalBytes)
        jobLog('transfer', `→ Target picked: ${tgt}`)
        toast('ok', i18n.t('offboarding.pickedTarget', { t: tgt }))
      }
      const r = await api.drive.copyBetweenUsers(p.source, tgt, folder, p.overwrite)
      patchJob('transfer', { result: r })
      const failed = Object.keys(r.failed || {}).length
      jobLog('transfer', `${r.canceled ? '⏹ Canceled' : '✓ Done'} — ${r.copied?.length || 0} copied, ${Object.keys(r.skipped || {}).length} skipped, ${failed} failed`)
      toast(r.canceled ? 'info' : 'ok', `${r.copied?.length || 0} copied${r.canceled ? ' (canceled)' : ''}`)
    } catch (e) {
      patchJob('transfer', { error: errMessage(e) })
      jobLog('transfer', `✗ Error: ${errMessage(e)}`)
    } finally {
      patchJob('transfer', { running: false, progress: '' })
    }
  }, [patchJob, jobLog, toast])

  const cancelTransfer = useCallback(() => {
    patchJob('transfer', { canceled: true, progress: 'Canceling…' })
    jobLog('transfer', '⏹ Cancel requested — stopping after the current file…')
    api.drive.cancelTransfer().catch(() => {})
  }, [patchJob, jobLog])

  const startCleanupScan = useCallback(async (p: CleanupParams) => {
    patchJob('cleanup', { ...emptyJob(), running: true, startedAt: Date.now(), progress: 'Starting…' })
    jobLog('cleanup', `▶ Scanning ${p.ownerType} · ${p.mode}…`)
    try {
      if (p.mode === 'duplicates') {
        const g = (await api.cleanup.findDuplicates(p.ownerType, p.ownerId)) || []
        patchJob('cleanup', { result: { mode: 'duplicates', groups: g } })
        jobLog('cleanup', `✓ Done — ${g.length} duplicate group(s)`)
      } else {
        const b = (await api.cleanup.findVersionBloat(p.ownerType, p.ownerId, 2, 3000)) || []
        patchJob('cleanup', { result: { mode: 'versions', bloat: b } })
        jobLog('cleanup', `✓ Done — ${b.length} file(s) with version bloat`)
      }
    } catch (e) {
      patchJob('cleanup', { error: errMessage(e) })
      jobLog('cleanup', `✗ Error: ${errMessage(e)}`)
    } finally {
      patchJob('cleanup', { running: false, progress: '' })
    }
  }, [patchJob, jobLog])

  const cancelCleanupScan = useCallback(() => {
    patchJob('cleanup', { canceled: true, progress: 'Canceling…' })
    jobLog('cleanup', '⏹ Cancel requested — returning partial results…')
    api.cleanup.cancelScan().catch(() => {})
  }, [patchJob, jobLog])

  // Bulk CSV runner: sequential on purpose (throttle-friendly; the Graph client
  // already retries 429s). Rows fail independently; cancel stops between rows.
  const bulkCancel = useRef(false)
  const startBulkRun = useCallback(async (items: BulkItem[]) => {
    bulkCancel.current = false
    patchJob('bulk', { ...emptyJob(), running: true, startedAt: Date.now(), progress: 'Starting…' })
    jobLog('bulk', `▶ Running ${items.length} row(s)…`)
    const results: BulkRowResult[] = []
    let ok = 0
    let failed = 0
    for (let i = 0; i < items.length; i++) {
      if (bulkCancel.current) {
        jobLog('bulk', `⏹ Canceled after ${i} of ${items.length} row(s).`)
        break
      }
      const it = items[i]
      patchJob('bulk', { progress: `${i + 1}/${items.length} — ${it.label}` })
      try {
        const d = await it.run()
        results.push({ label: it.label, ok: true, detail: d || undefined })
        ok++
      } catch (e) {
        const msg = errMessage(e)
        results.push({ label: it.label, ok: false, error: msg })
        failed++
        jobLog('bulk', `✗ ${it.label}: ${msg}`)
      }
    }
    jobLog('bulk', `✓ Done — ${ok} ok, ${failed} failed`)
    patchJob('bulk', { running: false, progress: '', result: results })
  }, [patchJob, jobLog])

  const cancelBulkRun = useCallback(() => {
    bulkCancel.current = true
    patchJob('bulk', { canceled: true, progress: 'Canceling…' })
    jobLog('bulk', '⏹ Cancel requested — stopping after the current row…')
  }, [patchJob, jobLog])

  // Playbook runs live here so their step report survives navigation, and the
  // completion toast fires even if the operator has left the page.
  const startPlaybook = useCallback(async (kind: 'onboard' | 'offboard', target: string, call: () => Promise<any>) => {
    patchJob('playbook', { ...emptyJob(), running: true, startedAt: Date.now(), progress: 'Starting…', steps: [] })
    jobLog('playbook', `▶ ${i18n.t(`playbooks.${kind}`)} — ${target}`)
    try {
      const r = await call()
      patchJob('playbook', { result: r, running: false, progress: '' })
      jobLog('playbook', r?.canceled ? '⏹ Canceled' : r?.ok ? '✓ Done' : '⚠ Finished with errors')
      toast(r?.canceled ? 'info' : r?.ok ? 'ok' : 'err',
        i18n.t(r?.canceled ? 'common.canceled' : r?.ok ? 'playbooks.doneToast' : 'playbooks.doneWithErrors'))
      return r
    } catch (e) {
      patchJob('playbook', { error: errMessage(e), running: false, progress: '' })
      jobLog('playbook', `✗ ${errMessage(e)}`)
      toast('err', errMessage(e))
      return null
    }
  }, [patchJob, jobLog, toast])

  const cancelPlaybook = useCallback(() => {
    patchJob('playbook', { canceled: true, progress: 'Canceling…' })
    jobLog('playbook', '⏹ Cancel requested — stopping after the current step…')
    api.playbooks.cancel().catch(() => {})
  }, [patchJob, jobLog])

  const value: Store = {
    status,
    setStatus,
    refreshStatus,
    connected: !!status?.connected,
    readOnly: !!status?.readOnly,
    domains,
    loadDomains,
    theme,
    toggleTheme: () => setTheme((t) => (t === 'dark' ? 'light' : 'dark')),
    accent,
    setAccent: setAccentState,
    safeMode,
    setSafeMode: (v) => {
      setSafeModeState(v)
      localStorage.setItem('safeMode', String(v))
    },
    access,
    hideUnavailable,
    setHideUnavailable: (v) => {
      setHideUnavailableState(v)
      localStorage.setItem('hideUnavailable', String(v))
    },
    checkAccess: async () => {
      try {
        const a = await api.access.probe()
        setAccess(a)
        localStorage.setItem('access', JSON.stringify(a))
      } catch {
        /* ignore */
      }
    },
    jobs,
    jobLog,
    patchJob,
    startTransfer,
    cancelTransfer,
    startCleanupScan,
    cancelCleanupScan,
    startBulkRun,
    cancelBulkRun,
    startPlaybook,
    cancelPlaybook,
    clearJob,
    cache,
    setCache,
    toasts,
    toast,
    dismiss: (id) => setToasts((t) => t.filter((x) => x.id !== id)),
  }

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>
}

export function useStore(): Store {
  const s = useContext(Ctx)
  if (!s) throw new Error('useStore must be used within StoreProvider')
  return s
}
