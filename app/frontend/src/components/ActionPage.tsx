import { useState, type ReactNode } from 'react'
import { X } from 'lucide-react'
import { Button } from './ui'

// Layout B: the result fills the page, actions live in a top toolbar, and each
// action opens a small slide-over drawer with just its form.

export interface Action {
  id: string
  label: string
  icon?: ReactNode
  variant?: 'primary' | 'subtle' | 'danger' | 'ghost'
  disabled?: boolean
  panel?: ReactNode // drawer form; if omitted, onClick runs immediately
  onClick?: () => void
}

export interface SearchBox {
  value: string
  onChange: (v: string) => void
  onSubmit?: () => void
  placeholder?: string
}

export function ActionPage({
  title,
  subtitle,
  search,
  actions,
  result,
}: {
  title: string
  subtitle?: string
  search?: SearchBox
  actions: Action[]
  result: ReactNode
}) {
  const [openId, setOpenId] = useState<string | null>(null)
  const open = actions.find((a) => a.id === openId)

  const trigger = (a: Action) => {
    if (a.panel) setOpenId((cur) => (cur === a.id ? null : a.id))
    a.onClick?.()
  }

  return (
    <div className="flex h-full flex-col">
      <header className="shrink-0 border-b border-[var(--border)] px-6 py-4">
        <h1 className="text-lg font-semibold">{title}</h1>
        {subtitle && <p className="mt-0.5 text-sm text-[var(--text-dim)]">{subtitle}</p>}
      </header>

      <div className="flex flex-wrap items-center gap-2 border-b border-[var(--border)] px-6 py-3">
        {search && (
          <form
            className="mr-1 min-w-[180px] flex-1 md:max-w-xs"
            onSubmit={(e) => {
              e.preventDefault()
              search.onSubmit?.()
            }}
          >
            <input
              value={search.value}
              onChange={(e) => search.onChange(e.target.value)}
              placeholder={search.placeholder || 'Search…'}
              className="w-full rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
            />
          </form>
        )}
        {actions.map((a) => (
          <Button
            key={a.id}
            variant={a.variant || 'subtle'}
            disabled={a.disabled}
            onClick={() => trigger(a)}
            className={openId === a.id ? 'ring-1 ring-[var(--accent)]' : ''}
          >
            {a.icon}
            {a.label}
          </Button>
        ))}
      </div>

      <div className="relative min-h-0 flex-1 overflow-hidden">
        <div className="h-full overflow-hidden p-4">
          <div className="h-full overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--bg-elev)]">
            {result}
          </div>
        </div>

        {open?.panel && (
          <>
            <div className="absolute inset-0 z-10 bg-black/30" onClick={() => setOpenId(null)} />
            <aside className="absolute right-0 top-0 z-20 flex h-full w-[340px] max-w-[86%] flex-col border-l border-[var(--border-strong)] bg-[var(--bg-elev)] shadow-2xl">
              <div className="flex items-center justify-between border-b border-[var(--border)] px-4 py-3">
                <h2 className="flex items-center gap-2 text-sm font-semibold">{open.icon}{open.label}</h2>
                <button onClick={() => setOpenId(null)} className="text-[var(--text-faint)] hover:text-[var(--text)]">
                  <X size={16} />
                </button>
              </div>
              <div className="min-h-0 flex-1 overflow-auto p-4">{open.panel}</div>
            </aside>
          </>
        )}
      </div>
    </div>
  )
}

// Field wrapper for drawer forms.
export function DrawerForm({ children }: { children: ReactNode }) {
  return <div className="flex flex-col gap-3">{children}</div>
}
