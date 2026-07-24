import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { RefreshCw, ChevronDown, ChevronRight, CheckCircle2, XCircle, PlayCircle, Loader2 } from 'lucide-react'
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
  const [resuming, setResuming] = useState('')

  const load = () => {
    api.journal.list(100).then((r) => { setRuns(r || []); setError(null) }).catch((e) => setError(errMessage(e)))
  }
  useEffect(load, [])

  const open = (id: string) => {
    if (openId === id) { setOpenId(''); setDetail(null); return }
    setOpenId(id); setDetail(null)
    api.journal.get(id).then(setDetail).catch((e) => toast('err', errMessage(e)))
  }

  const resume = (id: string) => {
    setResuming(id)
    api.drive.resumeCopy(id)
      .then((r) => {
        toast('ok', t('history.resumed', { n: r.copied?.length || 0 }))
        load()
        // Refresh the expanded detail too, so the events reflect the resume.
        if (openId === id) api.journal.get(id).then(setDetail).catch(() => {})
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
                {!detail && <Spinner />}
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
                      // Transfer journals carry item names only in the plan
                      // event; build an id→{name,size} map for the item rows.
                      const plan: Record<string, any> = {}
                      for (const ev of detail.events) {
                        if (ev.t === 'plan') {
                          for (const it of ((ev.data as any)?.items ?? [])) plan[it.id] = it
                        }
                      }
                      return detail.events.map((ev, i) => {
                        if (ev.t === 'item') {
                          const d: any = ev.data || {}
                          const name = plan[d.id]?.name ?? d.id
                          const ok = d.status === 'copied' || d.status === 'skipped'
                          return (
                            <div key={i} className="flex items-start gap-2 text-sm">
                              {ok ? <CheckCircle2 size={14} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={14} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                              <span className="truncate text-[var(--text)]">{name}</span>
                              <span className="shrink-0 text-xs text-[var(--text-faint)]">{d.status}</span>
                              {d.error && <span className="truncate text-xs text-[var(--danger)]">{d.error}</span>}
                            </div>
                          )
                        }
                        if (ev.t === 'step') {
                        const d: any = ev.data || {}
                        const name = d.nameKey ? t(d.nameKey, { defaultValue: d.name }) : d.name
                        return (
                          <div key={i} className="flex items-start gap-2 text-sm">
                            {d.ok ? <CheckCircle2 size={14} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={14} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                            <span className="text-[var(--text)]">{name}</span>
                            {d.detail && <span className="text-xs text-[var(--text-faint)]">{d.detail}</span>}
                            {d.error && <span className="text-xs text-[var(--danger)]">{d.error}</span>}
                          </div>
                        )
                      }
                        if (ev.t === 'log') {
                          return <div key={i} className="font-mono text-xs text-[var(--text-dim)]">{String((ev.data as any)?.text ?? '')}</div>
                        }
                        return null
                      })
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
    </Page>
  )
}
