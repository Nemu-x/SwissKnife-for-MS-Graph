import { useEffect, useRef, useState } from 'react'
import { ChevronDown, Check } from 'lucide-react'

export interface Option {
  value: string
  label: string
  sub?: string
}

// Checkbox dropdown for picking many options (licenses, groups, teams…).
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
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const h = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false) }
    document.addEventListener('mousedown', h)
    return () => document.removeEventListener('mousedown', h)
  }, [])

  const toggle = (v: string) =>
    onChange(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v])

  const label = selected.length === 0
    ? (placeholder || 'Select…')
    : `${selected.length} selected`

  return (
    <div className="relative" ref={ref}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
      >
        <span className={selected.length ? 'text-[var(--text)]' : 'text-[var(--text-faint)]'}>{label}</span>
        <ChevronDown size={15} className="text-[var(--text-faint)]" />
      </button>
      {open && (
        <div className="absolute z-30 mt-1 max-h-60 w-full overflow-auto rounded-lg border border-[var(--border-strong)] bg-[var(--bg-elev)] shadow-xl">
          {loading && <div className="px-3 py-2 text-xs text-[var(--text-faint)]">Loading…</div>}
          {!loading && options.length === 0 && <div className="px-3 py-2 text-xs text-[var(--text-faint)]">No options</div>}
          {options.map((o) => {
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
      )}
    </div>
  )
}
