import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { ChevronDown, Check } from 'lucide-react'

export interface Option {
  value: string
  label: string
  sub?: string
}

// Checkbox dropdown for picking many options. The panel renders in a portal with
// fixed positioning so it floats above the window and is never clipped by cards.
export function MultiSelect({
  options,
  selected,
  onChange,
  placeholder,
  loading,
}: {
  options: Option[]
  selected: string[]
  onChange: (v: string[]) => void
  placeholder?: string
  loading?: boolean
}) {
  const [open, setOpen] = useState(false)
  const [filter, setFilter] = useState('')
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

  const toggle = (v: string) =>
    onChange(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v])

  const label = selected.length === 0 ? (placeholder || 'Select…') : `${selected.length} selected`
  const shown = filter ? options.filter((o) => o.label.toLowerCase().includes(filter.toLowerCase())) : options

  return (
    <>
      <button
        ref={btnRef}
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
      >
        <span className={selected.length ? 'text-[var(--text)]' : 'text-[var(--text-faint)]'}>{label}</span>
        <ChevronDown size={15} className="text-[var(--text-faint)]" />
      </button>
      {open && rect && createPortal(
        <>
          <div className="fixed inset-0 z-[100]" onClick={() => { setOpen(false); setFilter('') }} />
          <div
            className="fixed z-[101] max-h-72 overflow-auto rounded-lg border border-[var(--border-strong)] bg-[var(--bg-elev)] shadow-2xl"
            style={{ top: rect.bottom + 4, left: rect.left, width: rect.width }}
          >
            <div className="sticky top-0 border-b border-[var(--border)] bg-[var(--bg-elev)] p-1.5">
              <input
                autoFocus
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
                placeholder="Search…"
                className="w-full rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
              />
            </div>
            {loading && <div className="px-3 py-2 text-xs text-[var(--text-faint)]">Loading…</div>}
            {!loading && shown.length === 0 && <div className="px-3 py-2 text-xs text-[var(--text-faint)]">No options</div>}
            {shown.map((o) => {
              const on = selected.includes(o.value)
              return (
                <button
                  key={o.value}
                  type="button"
                  onClick={() => toggle(o.value)}
                  className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-sm hover:bg-[var(--bg-elev-2)]"
                >
                  <span className={`flex h-4 w-4 shrink-0 items-center justify-center rounded border ${on ? 'border-[var(--accent)] bg-[var(--accent)] text-[var(--accent-fg)]' : 'border-[var(--border-strong)]'}`}>
                    {on && <Check size={12} />}
                  </span>
                  <span className="min-w-0">
                    <span className="block truncate text-[var(--text)]">{o.label}</span>
                    {o.sub && <span className="block truncate text-xs text-[var(--text-faint)]">{o.sub}</span>}
                  </span>
                </button>
              )
            })}
          </div>
        </>,
        document.body,
      )}
    </>
  )
}
