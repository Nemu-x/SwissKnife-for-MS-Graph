import { useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Play, FileUp, FileDown, CheckCircle2, XCircle } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Select, Spinner, Badge } from '../components/ui'
import { JobConsole } from '../components/JobConsole'
import { useConfirm } from '../lib/useConfirm'
import { useStore, type BulkItem, type BulkRowResult } from '../lib/store'
import { api } from '../lib/api'
import { parseCsv } from '../lib/csv'

type OpId = 'createUsers' | 'assignLicense' | 'addToGroup'

// Each operation: CSV columns (order-independent, matched by header name),
// which of them are required, and the API call for one row.
const OPS: Record<OpId, { headers: string[]; required: string[]; run: (r: Record<string, string>) => Promise<string | void> }> = {
  createUsers: {
    headers: ['displayName', 'upn', 'mailNickname', 'password', 'usageLocation'],
    required: ['displayName', 'upn', 'mailNickname', 'password'],
    run: async (r) => {
      await api.users.create(r.displayName, r.upn, r.mailNickname, r.password, true, r.usageLocation || '')
    },
  },
  assignLicense: {
    headers: ['upn', 'skuId'],
    required: ['upn', 'skuId'],
    run: async (r) => {
      await api.licensing.assign(r.upn, [r.skuId], [])
    },
  },
  addToGroup: {
    headers: ['upn', 'groupId'],
    required: ['upn', 'groupId'],
    run: async (r) => {
      await api.groups.addMember(r.groupId, r.upn)
    },
  },
}

export function BulkPage() {
  const { t } = useTranslation()
  const { readOnly, toast, jobs, startBulkRun, cancelBulkRun, clearJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const fileRef = useRef<HTMLInputElement>(null)

  const [op, setOp] = useState<OpId>('createUsers')
  const [csv, setCsv] = useState('')

  const job = jobs.bulk
  const busy = !!job?.running
  const results = (job?.result as BulkRowResult[] | null) || null

  const spec = OPS[op]

  // Parse + validate the CSV against the selected operation's headers.
  const parsed = (() => {
    if (!csv.trim()) return { rows: [] as Record<string, string>[], errors: [] as string[] }
    const raw = parseCsv(csv)
    if (raw.length < 2) return { rows: [], errors: raw.length ? [t('bulk.needData')] : [] }
    const header = raw[0].map((h) => h.trim())
    const unknown = header.filter((h) => !spec.headers.includes(h))
    const missing = spec.required.filter((h) => !header.includes(h))
    const errors: string[] = []
    if (unknown.length) errors.push(t('bulk.unknownColumns', { cols: unknown.join(', ') }))
    if (missing.length) errors.push(t('bulk.missingColumns', { cols: missing.join(', ') }))
    if (errors.length) return { rows: [], errors }
    const rows = raw.slice(1).map((cells) => {
      const rec: Record<string, string> = {}
      header.forEach((h, i) => { rec[h] = (cells[i] ?? '').trim() })
      return rec
    })
    rows.forEach((r, i) => {
      const bad = spec.required.filter((h) => !r[h])
      if (bad.length) errors.push(t('bulk.rowMissing', { n: i + 2, cols: bad.join(', ') }))
    })
    return { rows: errors.length ? [] : rows, errors }
  })()

  const openFile = (f: File | undefined) => {
    if (!f) return
    f.text().then(setCsv).catch(() => toast('err', 'read failed'))
  }

  const downloadTemplate = () => {
    const blob = new Blob([spec.headers.join(',') + '\n'], { type: 'text/csv' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `swissknife-${op}.csv`
    a.click()
    URL.revokeObjectURL(a.href)
  }

  const run = () => {
    const items: BulkItem[] = parsed.rows.map((r) => ({
      label: r.upn || r.displayName || JSON.stringify(r),
      run: () => spec.run(r),
    }))
    askConfirm('RUN', () => { startBulkRun(items) }, t('bulk.confirm', { n: items.length, op: t(`bulk.op.${op}`) }))
  }

  return (
    <Page title={t('bulk.title')} subtitle={t('bulk.subtitle')}>
      {confirmElement}
      <Card title={t('bulk.title')} className="mb-4">
        <div className="flex flex-wrap items-end gap-3">
          <label className="flex flex-col gap-1">
            <span className="text-xs text-[var(--text-dim)]">{t('bulk.operation')}</span>
            <Select value={op} onChange={(e) => setOp(e.target.value as OpId)} className="w-64">
              {(Object.keys(OPS) as OpId[]).map((k) => <option key={k} value={k}>{t(`bulk.op.${k}`)}</option>)}
            </Select>
          </label>
          <Button onClick={downloadTemplate}><FileDown size={15} /> {t('bulk.template')}</Button>
          <Button onClick={() => fileRef.current?.click()}><FileUp size={15} /> {t('bulk.openCsv')}</Button>
          <input ref={fileRef} type="file" accept=".csv,text/csv" className="hidden"
            onChange={(e) => { openFile(e.target.files?.[0]); e.target.value = '' }} />
          <Button variant="primary" className="ml-auto"
            disabled={readOnly || busy || parsed.rows.length === 0} onClick={run}>
            {busy ? <Spinner /> : <Play size={15} />} {t('bulk.run', { n: parsed.rows.length })}
          </Button>
        </div>
        <textarea
          value={csv} onChange={(e) => setCsv(e.target.value)}
          placeholder={spec.headers.join(',') + '\n…'}
          spellCheck={false}
          className="mt-3 h-40 w-full resize-y rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3 font-mono text-xs text-[var(--text)] outline-none focus:border-[var(--accent)]" />
        {parsed.errors.length > 0 && (
          <div className="mt-2 flex flex-col gap-1">
            {parsed.errors.slice(0, 6).map((e, i) => <div key={i} className="text-xs text-[var(--danger)]">{e}</div>)}
          </div>
        )}
        {parsed.rows.length > 0 && (
          <div className="mt-2 text-xs text-[var(--text-dim)]">{t('bulk.parsed', { n: parsed.rows.length })}</div>
        )}
      </Card>

      {job && (job.running || job.log.length > 0) && (
        <div className="mb-4"><JobConsole job={job} onCancel={cancelBulkRun} onClear={() => clearJob('bulk')} /></div>
      )}

      {results && results.length > 0 && (
        <Card title={t('bulk.results')}>
          <div className="mb-2 flex gap-2">
            <Badge kind="ok">{t('bulk.ok')}: {results.filter((r) => r.ok).length}</Badge>
            <Badge kind="danger">{t('bulk.failed')}: {results.filter((r) => !r.ok).length}</Badge>
          </div>
          <div className="flex max-h-96 flex-col gap-1 overflow-y-auto">
            {results.map((r, i) => (
              <div key={i} className="flex items-start gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm">
                {r.ok ? <CheckCircle2 size={15} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={15} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                <div className="min-w-0">
                  <span className="font-medium">{r.label}</span>
                  {r.error && <span className="ml-2 text-xs text-[var(--danger)]">{r.error}</span>}
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}
    </Page>
  )
}
