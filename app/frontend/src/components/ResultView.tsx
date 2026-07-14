import { useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Copy, Download } from 'lucide-react'
import { Button, Spinner, ErrorNote } from './ui'
import { useStore } from '../lib/store'
import { cellText, maskSensitive, pickColumns, toCSV, downloadText, type GraphRow } from '../lib/format'

type Tab = 'table' | 'json' | 'tree'

export function ResultView({
  data,
  loading,
  error,
}: {
  data: GraphRow[] | GraphRow | null
  loading?: boolean
  error?: string | null
}) {
  const { t } = useTranslation()
  const { safeMode, toast } = useStore()
  const [tab, setTab] = useState<Tab>('table')
  const [filter, setFilter] = useState('')

  const shown = useMemo(() => (safeMode ? maskSensitive(data) : data), [data, safeMode])
  const rows: GraphRow[] = useMemo(() => {
    if (Array.isArray(shown)) return shown
    if (shown && typeof shown === 'object') return [shown]
    return []
  }, [shown])

  const cols = useMemo(() => pickColumns(rows), [rows])
  const filtered = useMemo(() => {
    if (!filter) return rows
    const f = filter.toLowerCase()
    return rows.filter((r) => cols.some((c) => cellText(r[c]).toLowerCase().includes(f)))
  }, [rows, cols, filter])

  const copyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(shown, null, 2))
    toast('ok', 'JSON copied')
  }
  const exportCsv = () => {
    downloadText('export.csv', toCSV(cols, filtered))
    toast('ok', 'CSV exported')
  }

  return (
    <div className="flex h-full flex-col">
      <div className="flex items-center justify-between gap-2 border-b border-[var(--border)] px-3 py-2">
        <div className="flex gap-1">
          {(['table', 'json', 'tree'] as Tab[]).map((x) => (
            <button
              key={x}
              onClick={() => setTab(x)}
              className={`rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                tab === x ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]'
              }`}
            >
              {t(`common.${x}`)}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          {tab === 'table' && rows.length > 0 && (
            <>
              <input
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
                placeholder={t('common.search')}
                className="w-40 rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
              />
              <Button variant="ghost" onClick={exportCsv} className="!px-2 !py-1">
                <Download size={14} /> {t('common.exportCsv')}
              </Button>
            </>
          )}
          <Button variant="ghost" onClick={copyJson} className="!px-2 !py-1">
            <Copy size={14} /> JSON
          </Button>
        </div>
      </div>

      <div className="min-h-0 flex-1 overflow-auto">
        {loading && (
          <div className="flex h-full items-center justify-center gap-2 text-[var(--text-dim)]">
            <Spinner /> {t('common.loading')}
          </div>
        )}
        {!loading && error && <div className="p-4"><ErrorNote>{error}</ErrorNote></div>}
        {!loading && !error && rows.length === 0 && (
          <div className="flex h-full items-center justify-center text-sm text-[var(--text-faint)]">
            {t('common.empty')}
          </div>
        )}
        {!loading && !error && rows.length > 0 && tab === 'table' && <Table cols={cols} rows={filtered} />}
        {!loading && !error && rows.length > 0 && tab === 'json' && (
          <pre className="p-4 text-xs leading-relaxed text-[var(--text)]">{JSON.stringify(shown, null, 2)}</pre>
        )}
        {!loading && !error && rows.length > 0 && tab === 'tree' && <Tree value={shown} />}
      </div>
      {!loading && rows.length > 0 && (
        <footer className="border-t border-[var(--border)] px-3 py-1.5 text-xs text-[var(--text-faint)]">
          {t('common.rows', { n: filtered.length })}
        </footer>
      )}
    </div>
  )
}

function Table({ cols, rows }: { cols: string[]; rows: GraphRow[] }) {
  return (
    <table className="w-full border-collapse text-sm">
      <thead className="sticky top-0 bg-[var(--bg-elev-2)]">
        <tr>
          {cols.map((c) => (
            <th key={c} className="border-b border-[var(--border)] px-3 py-2 text-left font-semibold text-[var(--text-dim)]">
              {c}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((r, i) => (
          <tr key={i} className="hover:bg-[var(--bg-elev-2)]/50">
            {cols.map((c) => (
              <td key={c} className="max-w-xs truncate border-b border-[var(--border)] px-3 py-1.5" title={cellText(r[c])}>
                {cellText(r[c])}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function Tree({ value, k = '', depth = 0 }: { value: any; k?: string; depth?: number }) {
  const [open, setOpen] = useState(depth < 2)
  const isObj = value && typeof value === 'object'
  if (!isObj) {
    return (
      <div className="flex gap-2 px-4 py-0.5 text-xs" style={{ paddingLeft: 16 + depth * 14 }}>
        <span className="text-[var(--accent)]">{k}:</span>
        <span className="text-[var(--ok)]">{String(value)}</span>
      </div>
    )
  }
  const entries = Object.entries(value)
  return (
    <div>
      {k !== '' && (
        <div
          className="cursor-pointer px-4 py-0.5 text-xs text-[var(--text-dim)]"
          style={{ paddingLeft: 16 + depth * 14 }}
          onClick={() => setOpen((o) => !o)}
        >
          {open ? '▾' : '▸'} {k} {Array.isArray(value) ? `[${entries.length}]` : `{${entries.length}}`}
        </div>
      )}
      {open && entries.map(([kk, vv]) => <Tree key={kk} value={vv} k={kk} depth={depth + 1} />)}
    </div>
  )
}
