import { useEffect, useRef, useState, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { X, LayoutGrid, Table2, Rows2, Lock, Eraser } from 'lucide-react'
import { Spinner } from './ui'
import { useStore } from '../lib/store'
import type { PageId } from '../pages/registry'

// Action-first layout: everything the page can do is a visible tile, the tile
// expands into its form in place, and the raw data pane is a peer view instead
// of the whole screen. Which of the two leads is the operator's choice and is
// remembered per page.

export type PageView = 'actions' | 'data' | 'split'

export interface TaskAction {
  id: string
  label: string
  hint?: string // one line: what this action does, in the operator's words
  icon?: ReactNode
  variant?: 'primary' | 'danger'
  write?: boolean // needs write access — shows the read-only lock
  panel?: ReactNode // form; without one the tile just runs onClick
  // Caveats and prerequisites, shown next to the form instead of cluttering it.
  note?: ReactNode
  onClick?: () => void
}

// What the last run of an action did, shown on its tile. Toasts vanish after a
// few seconds; this stays until the next run.
export type ActionStatus = { ok: boolean; text: string; at: number }

export interface SearchBox {
  value: string
  onChange: (v: string) => void
  onSubmit?: () => void
  placeholder?: string
}

export function TaskPage({
  pageId,
  title,
  subtitle,
  search,
  actions,
  result,
  hasResult,
  status,
  busy,
  busyLabel,
  onClearResult,
}: {
  pageId: PageId
  title: string
  subtitle?: string
  search?: SearchBox
  actions: TaskAction[]
  result: ReactNode
  hasResult?: boolean
  status?: Record<string, ActionStatus>
  busy?: boolean // something is in flight — say so in every view mode
  busyLabel?: string // what exactly, when the page knows
  onClearResult?: () => void // drop the result and go back to the tiles
}) {
  const { t, i18n } = useTranslation()
  const { readOnly, access, pendingAction, requestAction } = useStore()
  const [openId, setOpenId] = useState<string | null>(null)
  const [view, setView] = useState<PageView>(
    () => (localStorage.getItem(`view.${pageId}`) as PageView) || 'split',
  )
  const setViewPersisted = (v: PageView) => {
    setView(v)
    localStorage.setItem(`view.${pageId}`, v)
  }

  // The task palette can ask for one of these actions by id (see tasks.ts).
  const actionsRef = useRef(actions)
  actionsRef.current = actions
  useEffect(() => {
    if (!pendingAction) return
    const a = actionsRef.current.find((x) => x.id === pendingAction)
    if (!a) return
    requestAction(null)
    if (a.panel) setOpenId(a.id)
    a.onClick?.()
  }, [pendingAction, requestAction])

  const trigger = (a: TaskAction) => {
    if (a.panel) setOpenId((cur) => (cur === a.id ? null : a.id))
    a.onClick?.()
  }

  // Nothing fetched yet? The data pane would only show "no data", so the tiles
  // take the whole page until there is something to look at.
  const showData = view !== 'actions' && !!hasResult
  const showActions = view !== 'data' || !hasResult
  const noAccess = pageId in access && access[pageId] === false

  const viewButton = (v: PageView, icon: ReactNode) => (
    <button
      key={v}
      onClick={() => setViewPersisted(v)}
      title={t(`view.${v}`)}
      aria-label={t(`view.${v}`)}
      aria-pressed={view === v}
      className={`flex items-center gap-1.5 rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
        view === v ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]'
      }`}
    >
      {icon}
      <span className="hidden sm:inline">{t(`view.${v}`)}</span>
    </button>
  )

  return (
    <div className="flex h-full flex-col">
      <header className="flex shrink-0 flex-wrap items-center justify-between gap-3 border-b border-[var(--border)] px-6 py-4">
        <div className="min-w-0">
          <h1 className="text-lg font-semibold">{title}</h1>
          {subtitle && <p className="mt-0.5 text-sm text-[var(--text-dim)]">{subtitle}</p>}
        </div>
        <div className="flex items-center gap-2">
          {/* Busy state lives in the header so it is visible in every view mode,
              including "actions" where the result pane is hidden. */}
          {busy && (
            <span className="flex items-center gap-2 text-xs text-[var(--accent2)]">
              <Spinner /> <span className="max-w-[280px] truncate">{busyLabel || t('common.working')}</span>
            </span>
          )}
          {hasResult && !busy && onClearResult && (
            <button
              onClick={onClearResult}
              className="flex items-center gap-1.5 rounded-md border border-[var(--border)] px-2 py-1 text-xs text-[var(--text-dim)] hover:border-[var(--accent)] hover:text-[var(--text)]"
            >
              <Eraser size={13} /> {t('view.clearResult')}
            </button>
          )}
          <div className="flex items-center gap-1 rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
            {viewButton('actions', <LayoutGrid size={13} />)}
            {viewButton('split', <Rows2 size={13} />)}
            {viewButton('data', <Table2 size={13} />)}
          </div>
        </div>
      </header>
      {busy && (
        <div className="h-0.5 shrink-0 overflow-hidden bg-[var(--bg-elev-2)]">
          <div className="h-full w-1/3 animate-pulse rounded-full bg-[var(--accent)]" />
        </div>
      )}

      {search && (
        <form
          className="shrink-0 border-b border-[var(--border)] px-6 py-3"
          onSubmit={(e) => { e.preventDefault(); search.onSubmit?.() }}
        >
          <input
            value={search.value}
            onChange={(e) => search.onChange(e.target.value)}
            placeholder={search.placeholder || t('common.search')}
            className="w-full max-w-md rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm outline-none focus:border-[var(--accent)]"
          />
        </form>
      )}

      <div className={`flex min-h-0 flex-1 flex-col ${showData && showActions ? '' : ''}`}>
        {showActions && (
          <div className={`min-h-0 overflow-auto p-4 ${showData ? 'max-h-[55%] shrink-0' : 'flex-1'}`}>
            {readOnly && (
              <p className="mb-3 flex items-center gap-1.5 text-xs text-[var(--warn)]">
                <Lock size={12} /> {t('safety.readOnlyOn')}
              </p>
            )}
            {noAccess && <p className="mb-3 text-xs text-[var(--warn)]">{t('view.noAccess')}</p>}
            <div className="grid gap-3 [grid-template-columns:repeat(auto-fill,minmax(232px,1fr))]">
              {actions.map((a) => {
                const locked = !!a.write && readOnly
                const st = status?.[a.id]
                const isOpen = openId === a.id
                return (
                  <div key={a.id} className="contents">
                    <button
                      onClick={() => trigger(a)}
                      className={`flex min-h-[86px] flex-col items-start gap-1 rounded-xl border p-3 text-left transition-colors
                        ${isOpen
                          ? 'border-[var(--accent)] bg-[var(--accent)]/10'
                          : 'border-[var(--border)] bg-[var(--bg-elev)] hover:border-[var(--border-strong)] hover:bg-[var(--bg-elev-2)]'}`}
                    >
                      <span className="flex w-full items-start gap-2">
                        <span className={`mt-0.5 shrink-0 ${a.variant === 'danger' ? 'text-[var(--danger)]' : 'text-[var(--accent)]'}`}>{a.icon}</span>
                        <span className="min-w-0 flex-1 text-sm font-medium leading-snug text-[var(--text)]">{a.label}</span>
                        {locked && <Lock size={12} className="mt-1 shrink-0 text-[var(--warn)]" />}
                      </span>
                      {a.hint && <span className="text-xs leading-snug text-[var(--text-faint)]">{a.hint}</span>}
                      {st && (
                        <span className={`mt-auto text-xs ${st.ok ? 'text-[var(--ok)]' : 'text-[var(--danger)]'}`}>
                          {st.text} · {new Date(st.at).toLocaleTimeString(i18n.language)}
                        </span>
                      )}
                    </button>
                    {/* The form opens in place, right under the tile's row. */}
                    {isOpen && a.panel && (
                      <section
                        style={{ gridColumn: '1 / -1' }}
                        className="rounded-xl border border-[var(--accent)]/40 bg-[var(--bg-elev)] p-4"
                      >
                        <div className="mb-3 flex items-center justify-between gap-2">
                          <h2 className="flex items-center gap-2 text-sm font-semibold">{a.icon}{a.label}</h2>
                          <button onClick={() => setOpenId(null)} aria-label={t('common.close')}
                            className="text-[var(--text-faint)] hover:text-[var(--text)]">
                            <X size={16} />
                          </button>
                        </div>
                        <div className="grid gap-5 md:grid-cols-[minmax(0,380px)_1fr]">
                          <div>{a.panel}</div>
                          {(a.note || a.hint) && (
                            <div className="max-w-prose border-t border-[var(--border)] pt-3 text-xs leading-relaxed text-[var(--text-faint)] md:border-l md:border-t-0 md:pl-5 md:pt-0">
                              {a.hint && <p className="text-[var(--text-dim)]">{a.hint}</p>}
                              {a.note && <div className="mt-2 flex flex-col gap-2">{a.note}</div>}
                            </div>
                          )}
                        </div>
                      </section>
                    )}
                  </div>
                )
              })}
            </div>
            {view === 'split' && !hasResult && (
              <p className="mt-4 text-xs text-[var(--text-faint)]">{t('view.dataHint')}</p>
            )}
          </div>
        )}

        {showData && (
          <div className="min-h-0 flex-1 overflow-hidden px-4 pb-4">
            <div className="h-full overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--bg-elev)]">
              {result}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

// Field wrapper for the forms inside tiles (same shape as the drawer forms).
export function TaskForm({ children }: { children: ReactNode }) {
  return <div className="flex flex-col gap-3">{children}</div>
}
