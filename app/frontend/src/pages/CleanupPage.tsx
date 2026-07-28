import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Trash2, Building2, Scissors, Files as FilesIcon, History } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { Button, Field, Select, Badge, Spinner, Input } from '../components/ui'
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
  const [sel, setSel] = useState<Record<string, boolean>>({})
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  // The scan runs in the store, so its progress/log survive page navigation.
  const job = jobs.cleanup
  const busy = !!job?.running
  const result = job?.result as { mode: Mode; groups?: services.DupGroup[]; bloat?: BloatRow[] } | null | undefined
  const groups = mode === 'duplicates' && result?.mode === 'duplicates' ? result.groups ?? null : null
  const bloat = mode === 'versions' && result?.mode === 'versions' ? result.bloat ?? null : null

  const scan = (m: Mode) => { setMode(m); setSel({}); startCleanupScan({ mode: m, ownerType, ownerId }) }

  const deleteExtras = () => {
    if (!groups) return
    const refs = groups.flatMap((g) => g.items.slice(1).map((i) => i.ref))
    if (refs.length === 0) return
    askConfirm('DELETE', async (c) => {
      try {
        const r = await api.cleanup.deleteItems(refs, c)
        toast('ok', `${r.deleted} deleted`)
        setStatus((s) => ({ ...s, duplicates: { ok: true, text: t('cleanup.deletedN', { n: r.deleted }), at: Date.now() } }))
        scan('duplicates')
      } catch (e) {
        // Leaving the previous green "deleted N" on the tile would contradict
        // the error the operator just saw.
        const m = errMessage(e)
        setStatus((s) => ({ ...s, duplicates: { ok: false, text: m, at: Date.now() } }))
        toast('err', m)
      }
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
        toast('ok', t('cleanup.trimmed', { n: r.removed }))
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
      patchJob('cleanup', { running: true, canceled: false, error: null, startedAt: Date.now(), progress: t('cleanup.trimming') })
      try {
        const rs = (await api.cleanup.trimVersionsMany(refs, keep, c)) || []
        const removed = rs.reduce((a, r) => a + r.removed, 0)
        const failed = rs.filter((r) => r.error).length
        // Tile and toast must tell the same story: a red tile saying "N trimmed"
        // next to a toast saying "N failures" is worse than either alone.
        const text = failed > 0 ? t('cleanup.trimFailures', { n: failed }) : t('cleanup.trimmed', { n: removed })
        toast(failed > 0 ? 'err' : 'ok', text)
        setStatus((s) => ({ ...s, versions: { ok: failed === 0, text, at: Date.now() } }))
        applyTrim(rs)
      } catch (e) {
        const m = errMessage(e)
        setStatus((s) => ({ ...s, versions: { ok: false, text: m, at: Date.now() } }))
        toast('err', m)
      }
      finally { patchJob('cleanup', { running: false, progress: '' }) }
    }, `${t('cleanup.trimSelected', { n: refs.length })} · ${humanBytes(selectedBytes)}`)
  }

  const totalDupWasted = (groups || []).reduce((a, g) => a + g.wasted, 0)
  const totalVerWasted = (bloat || []).reduce((a, b) => a + b.reclaimable, 0)

  // Both scans work against one drive: a user's OneDrive or a SharePoint site.
  const ownerFields = (
    <>
      <Select value={ownerType} onChange={(e) => { setOwnerType(e.target.value as any); setOwnerId('') }} className="w-full">
        <option value="user">{t('files.oneDrive')}</option>
        <option value="site">{t('files.sharePoint')}</option>
      </Select>
      <Field label={ownerType === 'user' ? t('common.user') : t('files.site')}>
        <EntityPicker value={ownerId} onChange={setOwnerId}
          load={ownerType === 'user' ? loadUsers : siteSizes ? loadSitesBySize : loadSites}
          reloadKey={`${ownerType}${siteSizes ? ':size' : ''}`}
          placeholder={ownerType === 'user' ? t('files.pickUser') : t('files.pickSite')} />
      </Field>
      {ownerType === 'site' && (
        <label className="flex items-center gap-2 text-xs text-[var(--text-dim)]">
          <input type="checkbox" checked={siteSizes} onChange={(e) => setSiteSizes(e.target.checked)} />
          {t('cleanup.showSiteSizes')}
        </label>
      )}
    </>
  )

  const actions: TaskAction[] = [
    {
      id: 'versions', label: t('cleanup.tileVersions'), hint: t('cleanup.hintVersions'),
      icon: <History size={16} />, variant: 'primary', write: true,
      note: <p>{t('cleanup.noteVersions')}</p>,
      onClick: () => setMode('versions'),
      panel: (
        <TaskForm>
          {ownerFields}
          <Field label={t('cleanup.keep')} hint={t('cleanup.keepHint')}>
            <Input type="number" value={keep} onChange={(e) => setKeep(Math.max(1, Number(e.target.value) || 1))} className="w-24" />
          </Field>
          <Button variant="primary" disabled={!ownerId || busy} onClick={() => scan('versions')}>
            {busy ? <Spinner /> : ownerType === 'site' ? <Building2 size={15} /> : <Search size={15} />}
            {busy ? t('cleanup.scanning') : t('cleanup.scanVersions')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'duplicates', label: t('cleanup.tileDuplicates'), hint: t('cleanup.hintDuplicates'),
      icon: <FilesIcon size={16} />, variant: 'primary', write: true,
      note: <p>{t('cleanup.noteDuplicates')}</p>,
      onClick: () => setMode('duplicates'),
      panel: (
        <TaskForm>
          {ownerFields}
          <Button variant="primary" disabled={!ownerId || busy} onClick={() => scan('duplicates')}>
            {busy ? <Spinner /> : ownerType === 'site' ? <Building2 size={15} /> : <Search size={15} />}
            {busy ? t('cleanup.scanning') : t('cleanup.scan')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  const resultPane = (
    <div className="flex h-full flex-col gap-3 overflow-auto p-3">
      {job && (job.log.length > 0 || job.running) && (
        <JobConsole job={job} onCancel={cancelCleanupScan} onClear={() => clearJob('cleanup')} />
      )}

      {mode === 'versions' && bloat && bloat.length === 0 && <p className="text-sm text-[var(--ok)]">{t('cleanup.noVersions')}</p>}
      {mode === 'versions' && bloat && bloat.length > 0 && (
        <>
          <div className="flex flex-wrap items-center gap-3">
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
              <div key={b.ref} className={`flex items-center gap-3 rounded-xl border border-[var(--border)] bg-[var(--bg)] p-3 ${b.trimmed ? 'opacity-60' : ''}`}>
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

      {mode === 'duplicates' && groups && groups.length === 0 && <p className="text-sm text-[var(--ok)]">{t('cleanup.noDupes')}</p>}
      {mode === 'duplicates' && groups && groups.length > 0 && (
        <>
          <div className="flex flex-wrap items-center gap-3">
            <span className="text-sm text-[var(--text-dim)]">{t('cleanup.totalWasted', { size: humanBytes(totalDupWasted) })}</span>
            <Button variant="danger" disabled={readOnly || busy} onClick={deleteExtras} className="ml-auto">
              <Trash2 size={15} /> {t('cleanup.deleteExtras')}
            </Button>
          </div>
          <div className="flex flex-col gap-2">
            {groups.map((g, i) => (
              <div key={i} className="rounded-xl border border-[var(--border)] bg-[var(--bg)] p-3">
                <div className="flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="truncate text-sm font-medium">{g.name}</div>
                    <div className="text-xs text-[var(--text-faint)]">{humanBytes(g.size)} · {t('cleanup.copies', { n: g.count })}</div>
                  </div>
                  <Badge kind="warn">{t('cleanup.wasted')}: {humanBytes(g.wasted)}</Badge>
                </div>
                <div className="mt-2 flex flex-col gap-1">
                  {g.items.map((it, j) => (
                    <div key={j} className="flex items-center gap-2 text-xs">
                      <span className={j === 0 ? 'text-[var(--ok)]' : 'text-[var(--text-faint)]'}>
                        {j === 0 ? t('cleanup.keepItem') : t('cleanup.extraItem')}
                      </span>
                      <span className="truncate text-[var(--text-dim)]">{it.path}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  )

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="cleanup"
        title={t('cleanup.title')}
        subtitle={t('cleanup.subtitle')}
        actions={actions}
        status={status}
        busy={busy}
        busyLabel={job?.progress || t('cleanup.scanning')}
        hasResult={!!job && (job.log.length > 0 || job.running || !!result)}
        onClearResult={() => clearJob('cleanup')}
        result={resultPane}
      />
    </>
  )
}
