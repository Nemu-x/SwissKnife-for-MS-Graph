import { createContext, useContext, useEffect, useState, type ReactNode, useCallback } from 'react'
import { api, errMessage, type Status } from './api'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import i18n from '../i18n'
import { applyAccent } from './color'

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
}
export type TransferParams = { source: string; target: string; dest: string; overwrite: boolean; usePool: boolean; pool: string }
export type CleanupParams = { mode: 'duplicates' | 'versions'; ownerType: 'user' | 'site'; ownerId: string }

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
  startTransfer: (p: TransferParams) => Promise<void>
  cancelTransfer: () => void
  startCleanupScan: (p: CleanupParams) => Promise<void>
  cancelCleanupScan: () => void
  clearJob: (key: string) => void

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

  // Clear cached domains on disconnect; loading is opt-in (checkbox on Connect).
  useEffect(() => {
    if (!status?.connected) setDomains([])
  }, [status?.connected])

  // Stream backend progress into the jobs, mounted once at the app root so it
  // keeps flowing regardless of which page (if any) is currently shown.
  useEffect(() => {
    const offP = EventsOn('transfer:progress', (d: any) => {
      const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
      patchJob('transfer', { progress: `${d.name} — ${pct}%` })
    })
    const offF = EventsOn('transfer:file', (d: any) => {
      const icon = d.status === 'copied' ? '✓' : d.status === 'skipped' ? '↷' : '✗'
      jobLog('transfer', `${icon} ${d.name}${d.reason ? ' — ' + d.reason : ''}`)
      patchJob('transfer', { progress: `${d.copied} copied…` })
    })
    const offC = EventsOn('cleanup:progress', (d: any) => {
      patchJob('cleanup', { progress: d.total > 0 ? `${d.stage} ${d.done}/${d.total}` : `${d.stage} ${d.done}…` })
    })
    const offCL = EventsOn('cleanup:log', (text: string) => jobLog('cleanup', text))
    return () => { offP(); offF(); offC(); offCL() }
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
    startTransfer,
    cancelTransfer,
    startCleanupScan,
    cancelCleanupScan,
    clearJob,
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
