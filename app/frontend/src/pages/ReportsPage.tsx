import { useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Download, BarChart3, Users, HardDrive, Mail, MessagesSquare, Building2 } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { Button, Field, Select, Spinner, ErrorNote } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { downloadText, humanBytes } from '../lib/format'

// Report ids are Graph's; their labels and the period labels are translated.
const REPORTS = ['office365ActiveUsers', 'oneDriveUsage', 'mailboxUsage', 'teamsUserActivity', 'sharePointUsage'] as const
type ReportId = typeof REPORTS[number]
const ICONS: Record<ReportId, JSX.Element> = {
  office365ActiveUsers: <Users size={16} />,
  oneDriveUsage: <HardDrive size={16} />,
  mailboxUsage: <Mail size={16} />,
  teamsUserActivity: <MessagesSquare size={16} />,
  sharePointUsage: <Building2 size={16} />,
}
const PERIODS = ['D7', 'D30', 'D90', 'D180']
const MAX_COLS = 8
const MAX_ROWS = 200

// Minimal CSV parser (handles quoted fields).
function parseCSV(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = [], field = '', q = false
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (q) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++ }
      else if (c === '"') q = false
      else field += c
    } else if (c === '"') q = true
    else if (c === ',') { row.push(field); field = '' }
    else if (c === '\n' || c === '\r') {
      if (field !== '' || row.length) { row.push(field); rows.push(row); row = []; field = '' }
      if (c === '\r' && text[i + 1] === '\n') i++
    } else field += c
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row) }
  return rows.filter((r) => r.length > 1)
}

interface Chart { label: string; unit: 'bytes' | 'num'; items: { name: string; value: number }[] }

// Picks a label column and the largest numeric column to visualize.
function buildChart(rows: string[][]): Chart | null {
  if (rows.length < 2) return null
  const header = rows[0]
  const body = rows.slice(1)
  const labelIdx = header.findIndex((h) => /name|display|url|site/i.test(h))
  let bestIdx = -1, bestSum = -1
  header.forEach((h, i) => {
    if (i === labelIdx) return
    const sum = body.reduce((a, r) => a + (Number(r[i]) || 0), 0)
    if (sum > bestSum && /byte|count|used|active|storage|message|meeting|call/i.test(h)) { bestSum = sum; bestIdx = i }
  })
  if (bestIdx < 0) return null
  const unit: Chart['unit'] = /byte|storage/i.test(header[bestIdx]) ? 'bytes' : 'num'
  const items = body
    .map((r) => ({ name: (labelIdx >= 0 ? r[labelIdx] : r[0]) || '—', value: Number(r[bestIdx]) || 0 }))
    .filter((x) => x.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 12)
  return { label: header[bestIdx], unit, items }
}

export function ReportsPage() {
  const { t } = useTranslation()
  const { toast, cache, setCache } = useStore()
  // Cache-backed: restore the last fetched CSV (and the params it was fetched
  // with) so returning to the page shows the table/chart without a re-fetch.
  const [report, setReport] = useState<ReportId>(() => cache['reports.params']?.report ?? 'oneDriveUsage')
  const [period, setPeriod] = useState<string>(() => cache['reports.params']?.period ?? 'D30')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [csv, setCsvLocal] = useState<string>(() => cache['reports.csv'] ?? '')
  const setCsv = (v: string) => { setCsvLocal(v); setCache('reports.csv', v) }
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  const table = useMemo(() => (csv ? parseCSV(csv) : []), [csv])
  const chart = useMemo(() => (table.length ? buildChart(table) : null), [table])
  const max = chart ? Math.max(...chart.items.map((i) => i.value), 1) : 1
  const fmt = (v: number) => (chart?.unit === 'bytes' ? humanBytes(v) : v.toLocaleString())

  const run = async (id: ReportId) => {
    setReport(id)
    setBusy(true); setError(null); setCsv('')
    try {
      const data = await api.reports.csv(id, period)
      setCsv(data)
      setCache('reports.params', { report: id, period })
      const rows = Math.max(0, parseCSV(data).length - 1)
      setStatus((s) => ({ ...s, [id]: { ok: true, text: t('reports.dataRows', { n: rows }), at: Date.now() } }))
      toast('ok', t('reports.fetched'))
    } catch (e) {
      const m = errMessage(e)
      setError(m)
      setStatus((s) => ({ ...s, [id]: { ok: false, text: m, at: Date.now() } }))
    } finally { setBusy(false) }
  }

  const periodField = (
    <Field label={t('reports.period')}>
      <Select value={period} onChange={(e) => setPeriod(e.target.value)} className="w-full">
        {PERIODS.map((p) => <option key={p} value={p}>{t('reports.lastDays', { n: p.slice(1) })}</option>)}
      </Select>
    </Field>
  )

  const actions: TaskAction[] = REPORTS.map((id) => ({
    id,
    label: t(`reports.name.${id}`),
    hint: t(`reports.hint.${id}`),
    icon: ICONS[id],
    variant: 'primary' as const,
    note: <p>{t('reports.noteCommon')}</p>,
    panel: (
      <TaskForm>
        {periodField}
        <Button variant="primary" disabled={busy} onClick={() => run(id)}>
          {busy ? <Spinner /> : <BarChart3 size={15} />} {t('common.run')}
        </Button>
        {csv && report === id && (
          <Button variant="subtle" onClick={() => downloadText(`${id}-${period}.csv`, csv)}>
            <Download size={15} /> {t('reports.download')}
          </Button>
        )}
        {error && <ErrorNote>{error}</ErrorNote>}
      </TaskForm>
    ),
  }))

  const resultPane = (
    <div className="flex h-full flex-col gap-4 overflow-auto p-3">
      {chart && (
        <div className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3">
          <div className="mb-2 text-sm font-medium">{t('reports.topChart', { label: chart.label, n: chart.items.length })}</div>
          <div className="flex flex-col gap-2.5">
            {chart.items.map((it, i) => (
              <div key={i} className="grid grid-cols-[minmax(120px,240px)_1fr_auto] items-center gap-3">
                <span className="truncate text-sm text-[var(--text-dim)]" title={it.name}>{it.name}</span>
                <div className="h-3 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                  <div className="h-full rounded bg-[var(--accent)]" style={{ width: `${(it.value / max) * 100}%` }} />
                </div>
                <span className="text-xs tabular-nums text-[var(--text)]">{fmt(it.value)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {table.length > 1 && (
        <div className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3">
          <div className="mb-2 flex items-center justify-between gap-3">
            <span className="text-sm font-medium">{t('reports.dataRows', { n: table.length - 1 })}</span>
            <Button variant="ghost" className="!px-2 !py-1 text-xs" onClick={() => downloadText(`${report}-${period}.csv`, csv)}>
              <Download size={14} /> {t('reports.download')}
            </Button>
          </div>
          {/* No silent truncation: say what is not on screen. */}
          {(table.length - 1 > MAX_ROWS || table[0].length > MAX_COLS) && (
            <p className="mb-2 text-xs text-[var(--warn)]">
              {t('reports.truncated', {
                rows: Math.min(table.length - 1, MAX_ROWS),
                totalRows: table.length - 1,
                cols: Math.min(table[0].length, MAX_COLS),
                totalCols: table[0].length,
              })}
            </p>
          )}
          <div className="max-h-[28rem] overflow-auto">
            <table className="w-full border-collapse text-xs">
              <thead className="sticky top-0 bg-[var(--bg-elev-2)]">
                <tr>
                  {table[0].slice(0, MAX_COLS).map((h, i) => (
                    <th key={i} className="border-b border-[var(--border)] px-2 py-1.5 text-left font-semibold text-[var(--text-dim)]">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {table.slice(1, MAX_ROWS + 1).map((r, ri) => (
                  <tr key={ri} className="hover:bg-[var(--bg-elev-2)]/50">
                    {r.slice(0, MAX_COLS).map((c, ci) => (
                      <td key={ci} className="max-w-[16rem] truncate border-b border-[var(--border)]/50 px-2 py-1" title={c}>{c}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )

  return (
    <TaskPage
      pageId="reports"
      title={t('reports.title')}
      subtitle={t('reports.subtitle')}
      actions={actions}
      status={status}
      busy={busy}
      busyLabel={t('reports.fetching')}
      hasResult={!!csv || busy || !!error}
      onClearResult={() => { setCsv(''); setError(null) }}
      result={resultPane}
    />
  )
}
