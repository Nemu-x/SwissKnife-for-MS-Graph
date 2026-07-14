import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Copy, ArrowRight } from 'lucide-react'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { Page } from '../components/Layout'
import { Card, Field, Input, Button, Badge, ErrorNote, Spinner } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

export function OffboardingPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()

  const [source, setSource] = useState('')
  const [target, setTarget] = useState('')
  const [dest, setDest] = useState('')
  const [overwrite, setOverwrite] = useState(false)

  const [preview, setPreview] = useState<services.CopyPreview | null>(null)
  const [previewing, setPreviewing] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const [copying, setCopying] = useState(false)
  const [progress, setProgress] = useState<{ name: string; pct: number } | null>(null)
  const [report, setReport] = useState<services.CopyResult | null>(null)

  useEffect(() => {
    const off = EventsOn('transfer:progress', (d: any) => {
      const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
      setProgress({ name: d.name, pct })
    })
    return () => off()
  }, [])

  const runPreview = async () => {
    setPreviewing(true); setError(null); setPreview(null); setReport(null)
    try {
      setPreview(await api.drive.offboardingPreview(source))
    } catch (e) { setError(errMessage(e)) } finally { setPreviewing(false) }
  }

  const runCopy = async () => {
    setCopying(true); setError(null); setReport(null); setProgress(null)
    try {
      const r = await api.drive.copyBetweenUsers(source, target, dest, overwrite)
      setReport(r)
      toast('ok', `${r.copied?.length || 0} ${t('offboarding.copied').toLowerCase()}`)
    } catch (e) { setError(errMessage(e)) } finally { setCopying(false); setProgress(null) }
  }

  return (
    <Page title={t('offboarding.title')} subtitle={t('offboarding.subtitle')}>
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-[minmax(340px,420px)_1fr]">
        <div className="flex flex-col gap-4">
          <Card title={t('offboarding.title')}>
            <div className="flex flex-col gap-3">
              <Field label={t('offboarding.source')}>
                <UpnInput value={source} onChange={setSource} placeholder="alice@contoso.com" />
              </Field>
              <div className="flex justify-center text-[var(--text-faint)]"><ArrowRight size={16} /></div>
              <Field label={t('offboarding.target')}>
                <UpnInput value={target} onChange={setTarget} placeholder="backup@contoso.com" />
              </Field>
              <Field label={t('offboarding.destFolder')} hint={t('offboarding.destHint')}>
                <Input value={dest} onChange={(e) => setDest(e.target.value)} placeholder="Backups/alice" />
              </Field>
              <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                <input type="checkbox" checked={overwrite} onChange={(e) => setOverwrite(e.target.checked)} />
                {t('offboarding.overwrite')}
              </label>

              <div className="flex gap-2">
                <Button variant="subtle" className="flex-1" disabled={!source || previewing} onClick={runPreview}>
                  {previewing ? <Spinner /> : <Search size={15} />} {t('offboarding.preview')}
                </Button>
                <Button variant="primary" className="flex-1" disabled={readOnly || !source || !target || copying} onClick={runCopy}>
                  {copying ? <Spinner /> : <Copy size={15} />} {copying ? t('offboarding.running') : t('offboarding.start')}
                </Button>
              </div>
              {error && <ErrorNote>{error}</ErrorNote>}
            </div>
          </Card>

          {preview && (
            <Card title={t('offboarding.preview')}>
              <p className="text-sm text-[var(--text-dim)]">
                {t('offboarding.previewResult', {
                  files: preview.files,
                  folders: preview.folders,
                  size: humanBytes(preview.totalBytes),
                })}
              </p>
            </Card>
          )}

          {progress && copying && (
            <div className="rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-3">
              <div className="mb-1 flex justify-between text-xs text-[var(--text-dim)]">
                <span className="truncate">{progress.name}</span><span>{progress.pct}%</span>
              </div>
              <div className="h-1.5 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                <div className="h-full bg-[var(--accent)] transition-all" style={{ width: `${progress.pct}%` }} />
              </div>
            </div>
          )}
        </div>

        <Card title={t('offboarding.report')} className="min-h-[300px]">
          {!report && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          {report && (
            <div className="flex flex-col gap-4">
              <div className="flex gap-2">
                <Badge kind="ok">{t('offboarding.copied')}: {report.copied?.length || 0}</Badge>
                <Badge kind="warn">{t('offboarding.skipped')}: {Object.keys(report.skipped || {}).length}</Badge>
                <Badge kind="danger">{t('offboarding.failed')}: {Object.keys(report.failed || {}).length}</Badge>
              </div>
              <ReportList title={t('offboarding.copied')} items={(report.copied || []).map((n) => [n, ''])} ok />
              <ReportList title={t('offboarding.skipped')} items={Object.entries(report.skipped || {})} />
              <ReportList title={t('offboarding.failed')} items={Object.entries(report.failed || {})} danger />
            </div>
          )}
        </Card>
      </div>
    </Page>
  )
}

function ReportList({ title, items, ok, danger }: { title: string; items: [string, string][]; ok?: boolean; danger?: boolean }) {
  if (items.length === 0) return null
  const color = danger ? 'var(--danger)' : ok ? 'var(--ok)' : 'var(--warn)'
  return (
    <div>
      <h3 className="mb-1 text-xs font-semibold" style={{ color }}>{title} ({items.length})</h3>
      <div className="max-h-48 overflow-auto rounded-lg border border-[var(--border)] bg-[var(--bg)]">
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
