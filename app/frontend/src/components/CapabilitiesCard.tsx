import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ChevronRight, RefreshCw } from 'lucide-react'
import { Card, Button, Input, Badge, Spinner } from './ui'
import { api, errMessage } from '../lib/api'
import { useStore } from '../lib/store'
import { entryLabel, type CatalogEntry } from './CatalogAction'

type Impl = { backend: string; state: 'runs' | 'ready' | 'unavailable'; via?: string; reason?: { key: string; params?: Record<string, string> } }
type Cap = { capability: string; action: string; page: string; danger: string; pack?: string; label?: Record<string, string>; impls: Impl[]; reason?: Impl['reason'] }

// What each capability runs through right now: the implementation that
// runs, the fallbacks behind it, and why the others cannot. Loaded on
// demand — checking PowerShell can take a few seconds.
export function CapabilitiesCard() {
  const { t } = useTranslation()
  // The report survives leaving Settings (checking takes seconds); the
  // store drops it when the connection changes.
  const { cache, setCache } = useStore()
  const caps = (cache['settings.capabilities'] as Cap[] | undefined) ?? null
  const setCaps = (v: Cap[]) => setCache('settings.capabilities', v)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [q, setQ] = useState('')
  const [onlyBlocked, setOnlyBlocked] = useState(false)

  // A recheck detects PowerShell again (a module may have been installed).
  const load = (recheck = false) => {
    setBusy(true)
    setError('')
    ;(recheck ? api.actions.refreshPowerShell().catch(() => {}) : Promise.resolve())
      .then(() => api.actions.capabilities())
      .then((v) => setCaps((v ?? []) as Cap[]))
      .catch((e) => setError(errMessage(e)))
      .finally(() => setBusy(false))
  }
  const reason = (r?: Impl['reason']) => r ? t(`actions.reasons.${r.key}`, { ...r.params, defaultValue: r.key }) : ''
  const backend = (b: string) => t(`actions.backends.${b}`, { defaultValue: b })
  const label = (c: Cap) => entryLabel({ id: c.action, label: c.label } as CatalogEntry, t)

  const needle = q.trim().toLowerCase()
  const shown = (caps ?? []).filter((c) => {
    if (onlyBlocked && !c.reason && c.impls.some((i) => i.state === 'runs')) return false
    return !needle || c.capability.toLowerCase().includes(needle) || label(c).toLowerCase().includes(needle)
  })
  const groups = new Map<string, Cap[]>()
  for (const c of shown) {
    const area = c.pack ? 'pack' : c.capability.split('.')[0]
    groups.set(area, [...(groups.get(area) ?? []), c])
  }
  const runnable = (caps ?? []).filter((c) => !c.reason && c.impls.some((i) => i.state === 'runs')).length

  return (
    <Card title={t('capabilities.title')} actions={caps && (
      <Button variant="subtle" onClick={() => load(true)} disabled={busy}>{busy ? <Spinner /> : <RefreshCw size={14} />} {t('capabilities.refresh')}</Button>
    )}>
      <div className="flex flex-col gap-3 p-4" data-capabilities>
        <p className="text-xs text-[var(--text-faint)]">{t('capabilities.intro')}</p>
        {!caps && (
          <div className="flex items-center gap-2">
            <Button variant="subtle" onClick={() => load()} disabled={busy}>{busy ? <Spinner /> : <ChevronRight size={14} />} {t('capabilities.show')}</Button>
            {error && <span className="text-xs text-[var(--danger)]">{error}</span>}
          </div>
        )}
        {caps && error && <p className="text-xs text-[var(--danger)]">{error}</p>}
        {caps && (
          <>
            <div className="flex flex-wrap items-center gap-3">
              <Input value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('capabilities.search')} className="min-w-0 flex-1" />
              <label className="flex items-center gap-1.5 text-xs text-[var(--text-dim)]">
                <input type="checkbox" checked={onlyBlocked} onChange={(e) => setOnlyBlocked(e.target.checked)} />
                {t('capabilities.onlyBlocked')}
              </label>
              <span className="text-xs text-[var(--text-faint)]">{t('capabilities.summary', { ok: runnable, total: caps.length })}</span>
            </div>
            {shown.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('capabilities.none')}</p>}
            {[...groups.entries()].map(([area, list]) => (
              <div key={area} className="flex flex-col gap-1">
                <h3 className="text-xs font-semibold uppercase tracking-wide text-[var(--text-faint)]">{t(`capabilities.area.${area}`, { defaultValue: area })}</h3>
                {list.map((c) => (
                  <div key={c.capability} data-capability={c.capability}
                    className="flex flex-col gap-1 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 sm:flex-row sm:items-center sm:gap-3">
                    <div className="min-w-0 sm:w-2/5">
                      <div className="truncate text-sm">{label(c)}</div>
                      <div className="truncate font-mono text-[11px] text-[var(--text-faint)]">{c.capability}</div>
                    </div>
                    <div className="flex min-w-0 flex-1 flex-col gap-1">
                      <div className="flex flex-wrap items-center gap-1">
                        {c.impls.map((i, n) => (
                          <span key={i.backend} className="flex items-center gap-1">
                            {n > 0 && <ChevronRight size={12} className="text-[var(--text-faint)]" />}
                            <span className={i.state === 'unavailable' ? 'line-through opacity-60' : ''}>
                              <Badge kind={i.state === 'runs' ? 'ok' : 'neutral'}>
                                {backend(i.backend)}
                                {i.via === 'worker' ? ` · ${t('capabilities.viaWorker')}` : ''}
                                {i.state === 'ready' ? ` · ${t('capabilities.fallback')}` : ''}
                              </Badge>
                            </span>
                          </span>
                        ))}
                      </div>
                      {c.reason && <span className="text-xs text-[var(--warn)]">{reason(c.reason)}</span>}
                      {/* Why each implementation cannot run, as text (not only on hover). */}
                      {c.impls.filter((i) => i.state === 'unavailable').map((i) => (
                        <span key={i.backend} className={`text-xs ${!c.reason && !c.impls.some((x) => x.state === 'runs') ? 'text-[var(--warn)]' : 'text-[var(--text-faint)]'}`}>
                          {backend(i.backend)}: {reason(i.reason)}
                        </span>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </>
        )}
      </div>
    </Card>
  )
}
