import { useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Download, BarChart3 } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Select, Spinner, ErrorNote } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { downloadText, humanBytes } from '../lib/format'

const REPORTS = [
  ['office365ActiveUsers', 'Office 365 active users'],
  ['oneDriveUsage', 'OneDrive usage'],
  ['mailboxUsage', 'Mailbox usage'],
  ['teamsUserActivity', 'Teams user activity'],
  ['sharePointUsage', 'SharePoint site usage'],
]
const PERIODS = ['D7', 'D30', 'D90', 'D180']

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
  const { toast } = useStore()
  const [report, setReport] = useState(REPORTS[1][0])
  const [period, setPeriod] = useState('D30')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [csv, setCsv] = useState('')

  const chart = useMemo(() => (csv ? buildChart(parseCSV(csv)) : null), [csv])
  const max = chart ? Math.max(...chart.items.map((i) => i.value), 1) : 1
  const fmt = (v: number) => (chart?.unit === 'bytes' ? humanBytes(v) : v.toLocaleString())

  const run = async () => {
    setBusy(true); setError(null); setCsv('')
    try {
      const data = await api.reports.csv(report, period)
      setCsv(data)
      toast('ok', t('reports.fetched'))
    } catch (e) { setError(errMessage(e)) } finally { setBusy(false) }
  }

  return (
    <Page title={t('reports.title')}>
      <div className="grid gap-4">
        <Card title={t('reports.title')}>
          <div className="flex flex-wrap items-end gap-3">
            <label className="flex flex-col gap-1">
              <span className="text-xs text-[var(--text-dim)]">{t('reports.report')}</span>
              <Select value={report} onChange={(e) => setReport(e.target.value)} className="w-64">
                {REPORTS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
              </Select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs text-[var(--text-dim)]">{t('reports.period')}</span>
              <Select value={period} onChange={(e) => setPeriod(e.target.value)}>
                {PERIODS.map((p) => <option key={p} value={p}>{p}</option>)}
              </Select>
            </label>
            <Button variant="primary" onClick={run} disabled={busy}>
              {busy ? <Spinner /> : <BarChart3 size={15} />} {t('common.run')}
            </Button>
            {csv && (
              <Button variant="ghost" onClick={() => downloadText(`${report}-${period}.csv`, csv)}>
                <Download size={15} /> {t('reports.download')}
              </Button>
            )}
          </div>
        </Card>

        {error && <ErrorNote>{error}</ErrorNote>}

        {chart && (
          <Card title={`${chart.label} — top ${chart.items.length}`}>
            <div className="flex flex-col gap-2.5">
              {chart.items.map((it, i) => (
                <div key={i} className="grid grid-cols-[minmax(120px,220px)_1fr_auto] items-center gap-3">
                  <span className="truncate text-sm text-[var(--text-dim)]" title={it.name}>{it.name}</span>
                  <div className="h-3 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                    <div className="h-full rounded bg-[var(--accent)]" style={{ width: `${(it.value / max) * 100}%` }} />
                  </div>
                  <span className="text-xs tabular-nums text-[var(--text)]">{fmt(it.value)}</span>
                </div>
              ))}
            </div>
          </Card>
        )}
      </div>
    </Page>
  )
}
