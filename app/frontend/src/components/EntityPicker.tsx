import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { ChevronDown, Loader2 } from 'lucide-react'
import type { Option } from './MultiSelect'
import { dropdownStyle } from '../lib/dropdown'

// Single-select combobox that loads its options on first open. You can also type
// a raw value (id/UPN) and press Enter to use it — so manual entry still works.
export function EntityPicker({
  value,
  onChange,
  load,
  placeholder,
  reloadKey,
}: {
  value: string
  onChange: (v: string) => void
  load: () => Promise<Option[]>
  placeholder?: string
  reloadKey?: string // change to force a reload (e.g. when a parent selection changes)
}) {
  const [open, setOpen] = useState(false)
  const [filter, setFilter] = useState('')
  const [options, setOptions] = useState<Option[] | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const btnRef = useRef<HTMLButtonElement>(null)
  const [rect, setRect] = useState<DOMRect | null>(null)

  const place = () => btnRef.current && setRect(btnRef.current.getBoundingClientRect())
  useLayoutEffect(() => { if (open) place() }, [open])
  useEffect(() => {
    if (!open) return
    const h = () => place()
    window.addEventListener('scroll', h, true)
    window.addEventListener('resize', h)
    return () => { window.removeEventListener('scroll', h, true); window.removeEventListener('resize', h) }
  }, [open])

  // Reset cached options when the reload key changes (parent context changed).
  useEffect(() => { setOptions(null) }, [reloadKey])

  const doLoad = async () => {
    setLoading(true); setError(null)
    try { setOptions(await load()) } catch (e: any) { setError(e?.message || String(e)); setOptions([]) } finally { setLoading(false) }
  }

  const openPanel = () => {
    setOpen(true)
    if (options === null && !loading) doLoad()
  }

  const selectedLabel = () => {
    if (!value) return placeholder || 'Select…'
    const hit = options?.find((o) => o.value === value)
    return hit ? hit.label : value
  }

  const shown = (options || []).filter((o) =>
    !filter || o.label.toLowerCase().includes(filter.toLowerCase()) || o.value.toLowerCase().includes(filter.toLowerCase()))

  const pick = (v: string) => { onChange(v); setOpen(false); setFilter('') }

  return (
    <>
      <button
        ref={btnRef}
        type="button"
        onClick={() => (open ? setOpen(false) : openPanel())}
        className="flex w-full items-center justify-between gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
      >
        <span className={`truncate ${value ? 'text-[var(--text)]' : 'text-[var(--text-faint)]'}`}>{selectedLabel()}</span>
        <ChevronDown size={15} className="shrink-0 text-[var(--text-faint)]" />
      </button>
      {open && rect && createPortal(
        <>
          <div className="fixed inset-0 z-[100]" onClick={() => { setOpen(false); setFilter('') }} />
          <div
            className="z-[101] overflow-auto rounded-lg border border-[var(--border-strong)] bg-[var(--bg-elev)] shadow-2xl"
            style={dropdownStyle(rect, { minWidth: 380, maxWidth: 560 })}
          >
            <div className="sticky top-0 border-b border-[var(--border)] bg-[var(--bg-elev)] p-1.5">
              <input
                autoFocus
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter' && filter.trim()) pick(filter.trim()) }}
                placeholder="Search or paste an ID…"
                className="w-full rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
              />
            </div>
            {loading && <div className="flex items-center gap-2 px-3 py-2 text-xs text-[var(--text-faint)]"><Loader2 size={13} className="animate-spin" /> Loading…</div>}
            {error && <div className="px-3 py-2 text-xs text-[var(--danger)]">{error}</div>}
            {!loading && !error && shown.length === 0 && (
              <div className="px-3 py-2 text-xs text-[var(--text-faint)]">
                {filter.trim() ? 'Press Enter to use typed value' : 'No options'}
              </div>
            )}
            {shown.map((o) => (
              <button
                key={o.value}
                type="button"
                onClick={() => pick(o.value)}
                className={`flex w-full flex-col items-start px-3 py-1.5 text-left hover:bg-[var(--bg-elev-2)] ${o.value === value ? 'bg-[var(--accent)]/10' : ''}`}
              >
                <span className="w-full truncate text-sm text-[var(--text)]">{o.label}</span>
                {o.sub && <span className="w-full truncate text-xs text-[var(--text-faint)]">{o.sub}</span>}
              </button>
            ))}
          </div>
        </>,
        document.body,
      )}
    </>
  )
}
