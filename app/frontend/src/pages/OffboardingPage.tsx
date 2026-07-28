import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Copy, HardDriveDownload, Users } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { Field, Input, Textarea, Button, Badge, ErrorNote, Spinner } from '../components/ui'
import { JobConsole } from '../components/JobConsole'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

export function OffboardingPage() {
  const { t } = useTranslation()
  const { readOnly, jobs, jobLog, startTransfer, cancelTransfer, clearJob } = useStore()

  const [source, setSource] = useState('')
  const [target, setTarget] = useState('')
  const [dest, setDest] = useState('')
  const [overwrite, setOverwrite] = useState(false)
  const [usePool, setUsePool] = useState(false)
  const [pool, setPool] = useState(localStorage.getItem('offboard.pool') || '')

  const [preview, setPreview] = useState<services.CopyPreview | null>(null)
  const [previewing, setPreviewing] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  // The copy runs in the store, so its state survives leaving and returning.
  const job = jobs.transfer
  const copying = !!job?.running
  const report = (job?.result as services.CopyResult | null) ?? null

  const runPreview = async () => {
    setPreviewing(true); setError(null); setPreview(null)
    jobLog('transfer', `🔍 Preview ${source}…`)
    try {
      const p = await api.drive.offboardingPreview(source)
      setPreview(p)
      const line = t('offboarding.previewResult', { files: p.files, folders: p.folders, size: humanBytes(p.totalBytes) })
      setStatus((s) => ({ ...s, preview: { ok: true, text: line, at: Date.now() } }))
      jobLog('transfer', `✓ Preview — ${p.files} files · ${p.folders} folders · ${humanBytes(p.totalBytes)}`)
    } catch (e) {
      const m = errMessage(e)
      setError(m)
      setStatus((s) => ({ ...s, preview: { ok: false, text: m, at: Date.now() } }))
      jobLog('transfer', `✗ Preview error: ${m}`)
    } finally { setPreviewing(false) }
  }

  const runCopy = () => {
    setError(null)
    startTransfer({ source, target, dest, overwrite, usePool, pool })
  }

  const sourceField = (
    <Field label={t('offboarding.source')}>
      <EntityPicker value={source} onChange={setSource} load={loadUsers} placeholder={t('users.pickUser')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'preview', label: t('offboarding.tilePreview'), hint: t('offboarding.hintPreview'),
      icon: <Search size={16} />, variant: 'primary',
      note: <p>{t('offboarding.notePreview')}</p>,
      panel: (
        <TaskForm>
          {sourceField}
          <Button variant="primary" disabled={!source || previewing} onClick={runPreview}>
            {previewing ? <Spinner /> : <Search size={15} />} {t('offboarding.preview')}
          </Button>
          {preview && (
            <p className="text-sm text-[var(--text-dim)]">
              {t('offboarding.previewResult', { files: preview.files, folders: preview.folders, size: humanBytes(preview.totalBytes) })}
            </p>
          )}
          {error && <ErrorNote>{error}</ErrorNote>}
        </TaskForm>
      ),
    },
    {
      id: 'copy', label: t('offboarding.tileCopy'), hint: t('offboarding.hintCopy'),
      icon: <HardDriveDownload size={16} />, variant: 'primary', write: true,
      note: (
        <>
          <p>{t('offboarding.noteCopy')}</p>
          {usePool && <p>{t('offboarding.notePool')}</p>}
        </>
      ),
      panel: (
        <TaskForm>
          {sourceField}
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={usePool} onChange={(e) => setUsePool(e.target.checked)} />
            {t('offboarding.usePool')}
          </label>
          {usePool ? (
            <Field label={t('offboarding.pool')}>
              <Textarea rows={3} value={pool} placeholder="backup1@contoso.com&#10;backup2@contoso.com"
                onChange={(e) => { setPool(e.target.value); localStorage.setItem('offboard.pool', e.target.value) }} />
            </Field>
          ) : (
            <Field label={t('offboarding.target')}>
              <EntityPicker value={target} onChange={setTarget} load={loadUsers} placeholder={t('users.pickUser')} />
            </Field>
          )}
          <Field label={t('offboarding.destFolder')} hint={t('offboarding.destHint')}>
            <Input value={dest} onChange={(e) => setDest(e.target.value)} />
          </Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={overwrite} onChange={(e) => setOverwrite(e.target.checked)} />
            {t('offboarding.overwrite')}
          </label>
          <Button variant="primary"
            disabled={readOnly || !source || (!usePool && !target) || (usePool && !pool.trim()) || copying}
            onClick={runCopy}>
            {copying ? <Spinner /> : <Copy size={15} />} {copying ? t('offboarding.running') : t('offboarding.start')}
          </Button>
          {error && <ErrorNote>{error}</ErrorNote>}
        </TaskForm>
      ),
    },
    {
      id: 'pool', label: t('offboarding.tilePool'), hint: t('offboarding.hintPool'),
      icon: <Users size={16} />,
      note: <p>{t('offboarding.notePool')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('offboarding.pool')}>
            <Textarea rows={5} value={pool} placeholder="backup1@contoso.com&#10;backup2@contoso.com"
              onChange={(e) => { setPool(e.target.value); localStorage.setItem('offboard.pool', e.target.value) }} />
          </Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={usePool} onChange={(e) => setUsePool(e.target.checked)} />
            {t('offboarding.usePool')}
          </label>
        </TaskForm>
      ),
    },
  ]

  const resultPane = (
    <div className="flex h-full flex-col gap-3 overflow-auto p-3">
      {job && (job.log.length > 0 || job.running)
        ? <JobConsole job={job} onCancel={cancelTransfer} onClear={() => clearJob('transfer')} />
        : <p className="text-sm text-[var(--text-faint)]">{t('offboarding.logEmpty')}</p>}

      {report && (
        <div className="flex flex-col gap-4 rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3">
          <div className="flex flex-wrap gap-2">
            <Badge kind="ok">{t('offboarding.copied')}: {report.copied?.length || 0}</Badge>
            <Badge kind="warn">{t('offboarding.skipped')}: {Object.keys(report.skipped || {}).length}</Badge>
            <Badge kind="danger">{t('offboarding.failed')}: {Object.keys(report.failed || {}).length}</Badge>
            {report.canceled && <Badge kind="warn">{t('common.canceled')}</Badge>}
          </div>
          <ReportList title={t('offboarding.copied')} items={(report.copied || []).map((n) => [n, ''])} ok />
          <ReportList title={t('offboarding.skipped')} items={Object.entries(report.skipped || {})} />
          <ReportList title={t('offboarding.failed')} items={Object.entries(report.failed || {})} danger />
        </div>
      )}
    </div>
  )

  return (
    <TaskPage
      pageId="offboarding"
      title={t('offboarding.title')}
      subtitle={t('offboarding.subtitle')}
      actions={actions}
      status={status}
      busy={copying || previewing}
      busyLabel={copying ? job?.progress || t('offboarding.running') : t('offboarding.preview')}
      hasResult={!!job && (job.log.length > 0 || job.running || !!report)}
      onClearResult={() => clearJob('transfer')}
      result={resultPane}
    />
  )
}

function ReportList({ title, items, ok, danger }: { title: string; items: [string, string][]; ok?: boolean; danger?: boolean }) {
  if (items.length === 0) return null
  const color = danger ? 'var(--danger)' : ok ? 'var(--ok)' : 'var(--warn)'
  return (
    <div>
      <h3 className="mb-1 text-xs font-semibold" style={{ color }}>{title} ({items.length})</h3>
      <div className="max-h-48 overflow-auto rounded-lg border border-[var(--border)] bg-[var(--bg-elev)]">
        {items.map(([name, reason], i) => (
          <div key={i} className="flex justify-between gap-3 border-b border-[var(--border)]/50 px-2.5 py-1 text-xs last:border-0">
            <span className="truncate text-[var(--text)]" title={name}>{name}</span>
            {reason && <span className="shrink-0 text-[var(--text-faint)]">{reason}</span>}
          </div>
        ))}
      </div>
    </div>
  )
}
