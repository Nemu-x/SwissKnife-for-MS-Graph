import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Trash2, Building2, Scissors, Files as FilesIcon, History } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Field, Select, Badge, Spinner, Input } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { JobConsole } from '../components/JobConsole'
import { loadUsers, loadSites, loadSitesBySize } from '../lib/pickers'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

type Mode = 'duplicates' | 'versions'

// A scan row plus client-side bookkeeping for items already trimmed in place.
type BloatRow = services.VersionBloat & { trimmed?: boolean; removedCount?: number }

export function CleanupPage() {
  const { t } = useTranslation()
  const { readOnly, toast, jobs, startCleanupScan, cancelCleanupScan, clearJob, patchJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()

  const [mode, setMode] = useState<Mode>('versions')
  const [ownerType, setOwnerType] = useState<'user' | 'site'>('user')
  const [ownerId, setOwnerId] = useState('')
  const [siteSizes, setSiteSizes] = useState(false)
  const [keep, setKeep] = useState(3)
  // Refs picked for a bulk trim; keyed by item ref so it survives re-sorts.
  const [sel, setSel] = useState<Record<string, boolean>>({})

  // The scan runs in the store, so its progress/log survive page navigation.
  const job = jobs.cleanup
  const busy = !!job?.running
  const result = job?.result as { mode: Mode; groups?: services.DupGroup[]; bloat?: BloatRow[] } | null | undefined
  // Only show results that match the currently selected tab.
  const groups = mode === 'duplicates' && result?.mode === 'duplicates' ? result.groups ?? null : null
  const bloat = mode === 'versions' && result?.mode === 'versions' ? result.bloat ?? null : null

  const scan = () => { setSel({}); startCleanupScan({ mode, ownerType, ownerId }) }

  const deleteExtras = () => {
    if (!groups) return
    const refs = groups.flatMap((g) => g.items.slice(1).map((i) => i.ref))
    if (refs.length === 0) return
    askConfirm('DELETE', async (c) => {
      try { const r = await api.cleanup.deleteItems(refs, c); toast('ok', `${r.deleted} deleted`); scan() }
      catch (e) { toast('err', errMessage(e)) }
    }, t('cleanup.deleteExtras'))
  }

  // Fold trim outcomes back into the scan result in place, so the list stays
  // useful without a full re-scan. Failed items keep their row untouched.
  const applyTrim = (results: { ref: string; removed: number; error?: string }[]) => {
    const byRef = new Map(results.map((r) => [r.ref, r]))
    const cur = job?.result as { mode: Mode; bloat?: BloatRow[] } | null | undefined
    if (!cur || cur.mode !== 'versions' || !cur.bloat) return
    const bloatNext = cur.bloat.map((b) => {
      const r = byRef.get(b.ref)
      if (!r || r.error) return b
      return { ...b, reclaimable: 0, versions: Math.min(b.versions, keep), trimmed: true, removedCount: r.removed }
    })
    patchJob('cleanup', { result: { ...cur, bloat: bloatNext } })
    setSel({})
  }

  const trim = (ref: string, name: string) =>
    askConfirm('TRIM', async (c) => {
      try {
        const r = await api.cleanup.trimVersions(ref, keep, c)
        toast('ok', `${r.removed} versions removed`)
        applyTrim([{ ref, removed: r.removed }])
      } catch (e) { toast('err', errMessage(e)) }
    }, `${t('cleanup.trim')} — ${name}`)

  const trimmable = (bloat || []).filter((b) => !b.trimmed)
  const selected = trimmable.filter((b) => sel[b.ref])
  const selectedBytes = selected.reduce((a, b) => a + b.reclaimable, 0)
  const allSelected = trimmable.length > 0 && selected.length === trimmable.length
  const toggleAll = () => {
    if (allSelected) { setSel({}); return }
    const next: Record<string, boolean> = {}
    trimmable.forEach((b) => { next[b.ref] = true })
    setSel(next)
  }

  const trimSelected = () => {
    const refs = selected.map((b) => b.ref)
    if (refs.length === 0) return
    askConfirm('TRIM', async (c) => {
      // Run as a job: the backend streams cleanup:progress/log events, and the
      // console's Cancel button (CancelScan) aborts between items.
      patchJob('cleanup', { running: true, canceled: false, error: null, startedAt: Date.now(), progress: 'Trimming…' })
      try {
        const rs = (await api.cleanup.trimVersionsMany(refs, keep, c)) || []
        const removed = rs.reduce((a, r) => a + r.removed, 0)
        const failed = rs.filter((r) => r.error).length
        toast(failed > 0 ? 'err' : 'ok', failed > 0 ? t('cleanup.trimFailures', { n: failed }) : `${removed} versions removed`)
        applyTrim(rs)
      } catch (e) { toast('err', errMessage(e)) }
      finally { patchJob('cleanup', { running: false, progress: '' }) }
    }, `${t('cleanup.trimSelected', { n: refs.length })} · ${humanBytes(selectedBytes)}`)
  }

  const totalDupWasted = (groups || []).reduce((a, g) => a + g.wasted, 0)
  const totalVerWasted = (bloat || []).reduce((a, b) => a + b.reclaimable, 0)

  return (
    <Page title={t('cleanup.title')} subtitle={t('cleanup.subtitle')}>
      {confirmElement}

      <div className="mb-4 inline-flex rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
        {([['versions', <History size={15} key="v" />, t('cleanup.tabVersions')], ['duplicates', <FilesIcon size={15} key="d" />, t('cleanup.tabDuplicates')]] as const).map(([m, icon, label]) => (
          <button key={m} onClick={() => setMode(m)}
            className={`flex items-center gap-2 rounded-md px-4 py-1.5 text-sm font-medium ${mode === m ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)]'}`}>
            {icon}{label}
          </button>
        ))}
      </div>

      <Card title={t('cleanup.title')} className="mb-4">
        <div className="flex flex-wrap items-end gap-3">
          <label className="flex flex-col gap-1">
            <span className="text-xs text-[var(--text-dim)]">Drive</span>
            <Select value={ownerType} onChange={(e) => { setOwnerType(e.target.value as any); setOwnerId('') }} className="w-48">
              <option value="user">OneDrive (user)</option>
              <option value="site">SharePoint (site)</option>
            </Select>
          </label>
          <Field label={ownerType === 'user' ? t('common.user') : 'SharePoint site'}>
            <div className="w-72">
              <EntityPicker value={ownerId} onChange={setOwnerId}
                load={ownerType === 'user' ? loadUsers : siteSizes ? loadSitesBySize : loadSites}
                reloadKey={`${ownerType}${siteSizes ? ':size' : ''}`}
                placeholder={ownerType === 'user' ? 'Pick a user…' : 'Pick a site…'} />
            </div>
          </Field>
          {mode === 'versions' && (
            <label className="flex flex-col gap-1">
              <span className="text-xs text-[var(--text-dim)]">{t('cleanup.keep')}</span>
              <Input type="number" value={keep} onChange={(e) => setKeep(Math.max(1, Number(e.target.value) || 1))} className="w-20" />
            </label>
          )}
          <Button variant="primary" disabled={!ownerId || busy} onClick={scan}>
            {busy ? <Spinner /> : ownerType === 'site' ? <Building2 size={15} /> : <Search size={15} />}
            {busy ? t('cleanup.scanning') : mode === 'versions' ? t('cleanup.scanVersions') : t('cleanup.scan')}
          </Button>
          {mode === 'duplicates' && groups && groups.length > 0 && (
            <Button variant="danger" disabled={readOnly} onClick={deleteExtras} className="ml-auto">
              <Trash2 size={15} /> {t('cleanup.deleteExtras')}
            </Button>
          )}
        </div>
        {ownerType === 'site' && (
          <label className="mt-3 flex w-fit items-center gap-2 border-t border-[var(--border)] pt-3 text-xs text-[var(--text-dim)]">
            <input type="checkbox" checked={siteSizes} onChange={(e) => setSiteSizes(e.target.checked)} />
            {t('cleanup.showSiteSizes')}
          </label>
        )}
      </Card>

      {job && (
        <div className="mb-4">
          <JobConsole job={job} onCancel={cancelCleanupScan} onClear={() => clearJob('cleanup')} />
        </div>
      )}

      {/* Versions */}
      {mode === 'versions' && bloat && bloat.length === 0 && <p className="text-sm text-[var(--ok)]">{t('cleanup.noVersions')}</p>}
      {mode === 'versions' && bloat && bloat.length > 0 && (
        <>
          <div className="mb-3 flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
              <input type="checkbox" checked={allSelected} disabled={trimmable.length === 0} onChange={toggleAll} />
              {t('cleanup.selectAll')}
            </label>
            <span className="text-sm text-[var(--text-dim)]">{t('cleanup.totalWasted', { size: humanBytes(totalVerWasted) })}</span>
            {selected.length > 0 && (
              <Button variant="danger" disabled={readOnly || busy} onClick={trimSelected} className="ml-auto">
                <Scissors size={15} /> {t('cleanup.trimSelected', { n: selected.length })} · {humanBytes(selectedBytes)}
              </Button>
            )}
          </div>
          <div className="flex flex-col gap-2">
            {bloat.map((b) => (
              <div key={b.ref} className={`flex items-center gap-3 rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] p-3 ${b.trimmed ? 'opacity-60' : ''}`}>
                <input type="checkbox" checked={!b.trimmed && !!sel[b.ref]} disabled={!!b.trimmed || busy}
                  onChange={(e) => setSel((s) => ({ ...s, [b.ref]: e.target.checked }))} />
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium" title={b.path}>{b.name}</div>
                  <div className="text-xs text-[var(--text-faint)]">{t('cleanup.versions', { n: b.versions })} · {humanBytes(b.currentSize)} {t('cleanup.current')}</div>
                </div>
                {b.trimmed
                  ? <Badge kind="ok">{t('cleanup.trimmed', { n: b.removedCount ?? 0 })}</Badge>
                  : <Badge kind="warn">{t('cleanup.wasted')}: {humanBytes(b.reclaimable)}</Badge>}
                <Button variant="danger" disabled={readOnly || busy || !!b.trimmed} onClick={() => trim(b.ref, b.name)}>
                  <Scissors size={14} /> {t('cleanup.trim')}
                </Button>
              </div>
            ))}
          </div>
        </>
      )}

      {/* Duplicates */}
      {mode === 'duplicates' && groups && groups.length === 0 && <p className="text-sm text-[var(--ok)]">{t('cleanup.noDupes')}</p>}
      {mode === 'duplicates' && groups && groups.length > 0 && (
        <>
          <div className="mb-3 text-sm text-[var(--text-dim)]">{t('cleanup.totalWasted', { size: humanBytes(totalDupWasted) })}</div>
          <div className="flex flex-col gap-2">
            {groups.map((g, i) => (
              <div key={i} className="rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] p-3">
                <div className="flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="truncate text-sm font-medium">{g.name}</div>
                    <div className="text-xs text-[var(--text-faint)]">{humanBytes(g.size)} each · {t('cleanup.copies', { n: g.count })}</div>
                  </div>
                  <Badge kind="warn">{t('cleanup.wasted')}: {humanBytes(g.wasted)}</Badge>
                </div>
                <div className="mt-2 flex flex-col gap-1">
                  {g.items.map((it, j) => (
                    <div key={j} className="flex items-center gap-2 text-xs">
                      <span className={j === 0 ? 'text-[var(--ok)]' : 'text-[var(--text-faint)]'}>{j === 0 ? '✓ keep' : '✗ extra'}</span>
                      <span className="truncate text-[var(--text-dim)]">{it.path}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </Page>
  )
}
