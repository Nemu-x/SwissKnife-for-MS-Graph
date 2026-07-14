import { useEffect, useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Copy, Download } from 'lucide-react'
import { Button, Spinner, ErrorNote } from './ui'
import { useStore } from '../lib/store'
import { cellText, maskSensitive, toCSV, downloadText, type GraphRow } from '../lib/format'

type Tab = 'table' | 'json' | 'tree'

// A value is "scalar" if it renders as a single readable cell.
function isScalar(v: any): boolean {
  return v == null || typeof v !== 'object'
}

// Primary label for a row in the master list.
function primaryLabel(row: GraphRow): string {
  for (const k of ['displayName', 'name', 'subject', 'topic', 'userPrincipalName', 'deviceName', 'id']) {
    if (row[k]) return cellText(row[k])
  }
  const first = Object.entries(row).find(([k, v]) => !k.startsWith('@') && isScalar(v))
  return first ? cellText(first[1]) : '(item)'
}

function secondaryLabel(row: GraphRow, primaryKey: string): string {
  for (const k of ['userPrincipalName', 'mail', 'skuPartNumber', 'id']) {
    if (k !== primaryKey && row[k]) return cellText(row[k])
  }
  return ''
}

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
  const [selected, setSelected] = useState(0)

  const shown = useMemo(() => (safeMode ? maskSensitive(data) : data), [data, safeMode])
  const rows: GraphRow[] = useMemo(() => {
    if (Array.isArray(shown)) return shown
    if (shown && typeof shown === 'object') return [shown]
    return []
  }, [shown])

  // Only scalar columns go into the CSV/compact list.
  const scalarCols = useMemo(() => {
    const seen: string[] = []
    for (const r of rows) {
      for (const [k, v] of Object.entries(r)) {
        if (!k.startsWith('@') && isScalar(v) && !seen.includes(k)) seen.push(k)
      }
    }
    return seen
  }, [rows])

  const filtered = useMemo(() => {
    if (!filter) return rows
    const f = filter.toLowerCase()
    return rows.filter((r) => Object.values(r).some((v) => cellText(v).toLowerCase().includes(f)))
  }, [rows, filter])

  useEffect(() => setSelected(0), [data])

  const current = filtered[Math.min(selected, filtered.length - 1)]

  const copyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(shown, null, 2))
    toast('ok', 'JSON copied')
  }
  const exportCsv = () => {
    downloadText('export.csv', toCSV(scalarCols, filtered))
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
                className="w-36 rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
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

      <div className="min-h-0 flex-1 overflow-hidden">
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
        {!loading && !error && rows.length > 0 && tab === 'table' && (
          <MasterDetail rows={filtered} selected={selected} onSelect={setSelected} current={current} />
        )}
        {!loading && !error && rows.length > 0 && tab === 'json' && (
          <pre className="h-full overflow-auto p-4 text-xs leading-relaxed text-[var(--text)]">{JSON.stringify(shown, null, 2)}</pre>
        )}
        {!loading && !error && rows.length > 0 && tab === 'tree' && (
          <div className="h-full overflow-auto"><Tree value={shown} /></div>
        )}
      </div>
      {!loading && rows.length > 0 && (
        <footer className="border-t border-[var(--border)] px-3 py-1.5 text-xs text-[var(--text-faint)]">
          {t('common.rows', { n: filtered.length })}
        </footer>
      )}
    </div>
  )
}

// Master (compact list) on the left, readable key/value detail on the right.
function MasterDetail({
  rows,
  selected,
  onSelect,
  current,
}: {
  rows: GraphRow[]
  selected: number
  onSelect: (i: number) => void
  current: GraphRow | undefined
}) {
  const single = rows.length === 1
  return (
    <div className={`grid h-full ${single ? 'grid-cols-1' : 'grid-cols-[minmax(180px,260px)_1fr]'}`}>
      {!single && (
        <div className="overflow-auto border-r border-[var(--border)]">
          {rows.map((r, i) => {
            const primary = primaryLabel(r)
            const secondary = secondaryLabel(r, 'displayName')
            return (
              <button
                key={i}
                onClick={() => onSelect(i)}
                className={`flex w-full flex-col items-start gap-0.5 border-b border-[var(--border)] px-3 py-2 text-left transition-colors
                  ${i === selected ? 'bg-[var(--accent)]/12' : 'hover:bg-[var(--bg-elev-2)]'}`}
              >
                <span className="truncate text-sm font-medium text-[var(--text)]">{primary}</span>
                {secondary && <span className="truncate text-xs text-[var(--text-faint)]">{secondary}</span>}
              </button>
            )
          })}
        </div>
      )}
      <div className="overflow-auto p-4">
        {current ? <Detail row={current} /> : <p className="text-sm text-[var(--text-faint)]">—</p>}
      </div>
    </div>
  )
}

function Detail({ row }: { row: GraphRow }) {
  const entries = Object.entries(row).filter(([k]) => !k.startsWith('@odata'))
  return (
    <dl className="flex flex-col gap-1.5">
      {entries.map(([k, v]) => (
        <div key={k} className="grid grid-cols-[minmax(120px,200px)_1fr] gap-3 border-b border-[var(--border)]/60 py-1">
          <dt className="truncate text-xs font-medium text-[var(--text-dim)]" title={k}>{k}</dt>
          <dd className="min-w-0 text-sm text-[var(--text)]">
            {isScalar(v) ? (
              <span className="break-words">{cellText(v)}</span>
            ) : (
              <Expandable value={v} />
            )}
          </dd>
        </div>
      ))}
    </dl>
  )
}

function Expandable({ value }: { value: any }) {
  const [open, setOpen] = useState(false)
  const count = Array.isArray(value) ? value.length : Object.keys(value).length
  const label = Array.isArray(value) ? `array [${count}]` : `object {${count}}`
  return (
    <div>
      <button onClick={() => setOpen((o) => !o)} className="text-xs text-[var(--accent2)] hover:underline">
        {open ? '▾' : '▸'} {label}
      </button>
      {open && (
        <pre className="mt-1 max-h-64 overflow-auto rounded-md bg-[var(--bg)] p-2 text-xs text-[var(--text-dim)]">
          {JSON.stringify(value, null, 2)}
        </pre>
      )}
    </div>
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
