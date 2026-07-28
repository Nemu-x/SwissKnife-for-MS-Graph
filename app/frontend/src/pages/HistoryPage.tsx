import { useEffect, useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { RefreshCw, ChevronDown, ChevronRight, CheckCircle2, XCircle, PlayCircle, Loader2, Clock } from 'lucide-react'
import { Page } from '../components/Layout'
import { Button, Badge, Spinner, ErrorNote } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import type { journal } from '../../wailsjs/go/models'

// History reads the persistent run journal: every playbook/transfer run
// survives app restarts here; interrupted cloud copies offer Resume.
export function HistoryPage() {
  const { t, i18n } = useTranslation()
  const { connected, toast } = useStore()
  const [runs, setRuns] = useState<journal.RunSummary[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [openId, setOpenId] = useState('')
  const [detail, setDetail] = useState<journal.Run | null>(null)
  const [detailErr, setDetailErr] = useState('')
  const [resuming, setResuming] = useState('')
  const [tab, setTab] = useState<'runs' | 'actions'>('runs')

  const load = () => {
    api.journal.list(100).then((r) => { setRuns(r || []); setError(null) }).catch((e) => setError(errMessage(e)))
  }
  useEffect(load, [])

  // Stale-response protection: async callbacks must consult CURRENT state
  // (refs), never their closure — any open/close/resume bumps the request
  // generation and only the matching response may land.
  const detailRequest = useRef(0)
  const openIdRef = useRef('')
  useEffect(() => { openIdRef.current = openId }, [openId])

  const fetchDetail = (id: string) => {
    const request = ++detailRequest.current
    setDetail(null) // a failed refresh must show the error, never stale content
    setDetailErr('')
    api.journal.get(id)
      .then((r) => { if (detailRequest.current === request) setDetail(r) })
      .catch((e) => { if (detailRequest.current === request) setDetailErr(errMessage(e)) })
  }

  const open = (id: string) => {
    if (openId === id) {
      ++detailRequest.current // invalidate any in-flight fetch for this panel
      openIdRef.current = '' // sync: async callbacks must see the change NOW
      setOpenId(''); setDetail(null); setDetailErr('')
      return
    }
    openIdRef.current = id
    setOpenId(id)
    fetchDetail(id)
  }

  const resume = (id: string) => {
    setResuming(id)
    api.drive.resumeCopy(id)
      .then((r) => {
        toast('ok', t('history.resumed', { n: r.copied?.length || 0 }))
        load()
        // Refresh the detail only if THIS run is still the open one (current
        // state via ref, not the closure) — failures surface like in open().
        if (openIdRef.current === id) fetchDetail(id)
      })
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setResuming(''))
  }

  const fmtTime = (s: string) => new Date(s).toLocaleString(i18n.language)

  const badge = (r: journal.RunSummary) => {
    if (!r.endedAt && !r.interrupted) return <Badge kind="warn">{t('history.running')}</Badge>
    if (r.interrupted) return <Badge kind="warn">{t('history.interrupted')}</Badge>
    if (r.summary?.canceled) return <Badge kind="warn">{t('common.canceled')}</Badge>
    if ((r.summary?.failed ?? 0) > 0 || r.summary?.ok === false) return <Badge kind="danger">{t('history.failed')}</Badge>
    return <Badge kind="ok">{t('history.ok')}</Badge>
  }

  const kindLabel = (r: journal.RunSummary) =>
    r.kind === 'playbook' ? t('nav.playbooks') : r.kind === 'transfer' ? t('offboarding.title') : r.kind

  return (
    <Page title={t('history.title')} subtitle={t('history.subtitle')}>
      {/* Two journals, one page: multi-step runs and the flat write log. */}
      <div className="mb-4 inline-flex rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
        {([['runs', t('history.tabRuns')], ['actions', t('history.tabActions')]] as const).map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)}
            className={`rounded-md px-4 py-1.5 text-sm font-medium ${tab === id ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)]'}`}>
            {label}
          </button>
        ))}
      </div>

      {tab === 'actions' && <ActionLog />}

      {tab === 'runs' && <>
      <div className="mb-3 flex items-center gap-2">
        <Button variant="subtle" onClick={load}><RefreshCw size={15} /> {t('common.refresh')}</Button>
      </div>
      {error && <ErrorNote>{error}</ErrorNote>}
      {runs && runs.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('history.empty')}</p>}
      {!runs && !error && <Spinner />}
      <div className="flex flex-col gap-2">
        {runs?.map((r) => (
          <div key={r.opId} className="rounded-xl border border-[var(--border)] bg-[var(--bg-elev)]">
            <button onClick={() => open(r.opId)} className="flex w-full items-center gap-3 px-4 py-3 text-left">
              {openId === r.opId ? <ChevronDown size={15} className="shrink-0" /> : <ChevronRight size={15} className="shrink-0" />}
              <span className="w-28 shrink-0 text-xs font-medium text-[var(--text-dim)]">{kindLabel(r)}</span>
              <span className="min-w-0 flex-1 truncate text-sm text-[var(--text)]">{r.target || r.opId}</span>
              <span className="shrink-0 text-xs text-[var(--text-faint)]">{fmtTime(String(r.startedAt))}</span>
              {badge(r)}
            </button>
            {openId === r.opId && (
              <div className="border-t border-[var(--border)] px-4 py-3">
                {!detail && detailErr && <ErrorNote>{detailErr}</ErrorNote>}
                {!detail && !detailErr && <Spinner />}
                {detail && (
                  <div className="flex flex-col gap-1.5">
                    {r.kind === 'transfer' && r.interrupted && (
                      <div className="mb-2">
                        {connected ? (
                          <Button variant="primary" disabled={!!resuming} onClick={() => resume(r.opId)}>
                            {resuming === r.opId ? <Loader2 size={15} className="animate-spin" /> : <PlayCircle size={15} />} {t('history.resume')}
                          </Button>
                        ) : (
                          <p className="text-xs text-[var(--warn)]">{t('history.connectToResume')}</p>
                        )}
                      </div>
                    )}
                    {(() => {
                      // Transfer journals log several events per item (inflight,
                      // then the outcome) and names live only in the plan event —
                      // aggregate to one row per item with its LAST status.
                      const plan: Record<string, any> = {}
                      const itemOrder: string[] = []
                      const itemLast: Record<string, any> = {}
                      for (const ev of detail.events) {
                        if (ev.t === 'plan') {
                          for (const it of ((ev.data as any)?.items ?? [])) plan[it.id] = it
                        } else if (ev.t === 'item') {
                          const d: any = ev.data || {}
                          if (!(d.id in itemLast)) itemOrder.push(d.id)
                          itemLast[d.id] = d
                        }
                      }
                      const rows = detail.events.map((ev, i) => {
                        if (ev.t === 'step') {
                          const d: any = ev.data || {}
                          const name = d.nameKey ? t(d.nameKey, { defaultValue: d.name }) : d.name
                          return (
                            <div key={`s${i}`} className="flex items-start gap-2 text-sm">
                              {d.ok ? <CheckCircle2 size={14} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={14} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                              <span className="text-[var(--text)]">{name}</span>
                              {d.detail && <span className="text-xs text-[var(--text-faint)]">{d.detail}</span>}
                              {d.error && <span className="text-xs text-[var(--danger)]">{d.error}</span>}
                            </div>
                          )
                        }
                        if (ev.t === 'log') {
                          return <div key={`l${i}`} className="font-mono text-xs text-[var(--text-dim)]">{String((ev.data as any)?.text ?? '')}</div>
                        }
                        return null
                      })
                      const itemRows = itemOrder.map((id) => {
                        const d = itemLast[id]
                        const name = plan[id]?.name ?? id
                        // inflight/pending are transient, not failures: amber clock.
                        const icon = d.status === 'copied' || d.status === 'skipped'
                          ? <CheckCircle2 size={14} className="mt-0.5 shrink-0 text-[var(--ok)]" />
                          : d.status === 'failed'
                            ? <XCircle size={14} className="mt-0.5 shrink-0 text-[var(--danger)]" />
                            : <Clock size={14} className="mt-0.5 shrink-0 text-[var(--warn)]" />
                        return (
                          <div key={`i${id}`} className="flex items-start gap-2 text-sm">
                            {icon}
                            <span className="truncate text-[var(--text)]">{name}</span>
                            <span className="shrink-0 text-xs text-[var(--text-faint)]">{d.status}</span>
                            {d.error && <span className="truncate text-xs text-[var(--danger)]">{d.error}</span>}
                          </div>
                        )
                      })
                      return [...rows, ...itemRows]
                    })()}
                    {detail.summary && (
                      <p className="mt-1 text-xs text-[var(--text-faint)]">
                        {t('history.summaryLine', {
                          copied: detail.summary.copied ?? detail.summary.steps ?? 0,
                          skipped: detail.summary.skipped ?? 0,
                          failed: detail.summary.failed ?? 0,
                        })}
                      </p>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
      </>}
    </Page>
  )
}

// The local write/destructive log (ADR-002): what this app did on this machine,
// as opposed to the tenant-side audit under Insights.
function ActionLog() {
  const { t } = useTranslation()
  const [entries, setEntries] = useState<any[] | null>(null)
  const [loading, setLoading] = useState(true)

  const load = () => {
    setLoading(true)
    api.audit.activity(200)
      .then((e) => setEntries((e || []).reverse()))
      .catch(() => setEntries([]))
      .finally(() => setLoading(false))
  }
  useEffect(() => { load() }, [])

  // Playbook summaries are stored as {key, params} JSON so they render in the
  // UI language; any other detail stays as-is.
  const detailText = (d: string) => {
    if (d && d.startsWith('{')) {
      try {
        const o = JSON.parse(d)
        if (o.key) return String(t(o.key, { ...o.params, defaultValue: d }))
      } catch { /* plain detail */ }
    }
    return d
  }

  return (
    <>
      <div className="mb-3 flex items-center gap-2">
        <Button variant="subtle" onClick={load} disabled={loading}>
          {loading ? <Spinner /> : <RefreshCw size={15} />} {t('common.refresh')}
        </Button>
        <span className="text-xs text-[var(--text-faint)]">{t('activity.subtitle')}</span>
      </div>
      {loading && <Spinner />}
      {!loading && entries?.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
      <div className="flex flex-col gap-1">
        {(entries || []).map((e, i) => (
          <div key={i} className="flex items-center gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] px-3 py-2 text-sm">
            {e.ok ? <CheckCircle2 size={15} className="shrink-0 text-[var(--ok)]" /> : <XCircle size={15} className="shrink-0 text-[var(--danger)]" />}
            <span className="shrink-0 font-mono text-xs text-[var(--accent)]">{e.action}</span>
            <span className="truncate text-[var(--text-dim)]">{e.target}</span>
            {e.detail && <span className="truncate text-xs text-[var(--text-faint)]">{detailText(e.detail)}</span>}
            <span className="ml-auto shrink-0 text-xs text-[var(--text-faint)]">{new Date(e.time).toLocaleString()}</span>
          </div>
        ))}
      </div>
    </>
  )
}
