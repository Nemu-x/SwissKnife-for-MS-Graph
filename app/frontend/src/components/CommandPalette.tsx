import { useEffect, useMemo, useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, CornerDownLeft, Lock } from 'lucide-react'
import { LOCAL_PAGES } from './Layout'
import { useStore } from '../lib/store'
import { TASKS, TASK_GROUPS, type Task } from '../lib/tasks'
import type { PageId } from '../pages/registry'

// "What do you need to do?" — the task-first entry point into the app.
// Everything here is a job, not an endpoint; picking one navigates to the page
// that does it and opens the right drawer via the store handshake.
export function CommandPalette({
  open,
  onClose,
  onNavigate,
}: {
  open: boolean
  onClose: () => void
  onNavigate: (p: PageId) => void
}) {
  const { t } = useTranslation()
  const { connected, readOnly, access, hideUnavailable, requestAction } = useStore()
  const [q, setQ] = useState('')
  const [active, setActive] = useState(0)
  const listRef = useRef<HTMLDivElement>(null)
  // Opening under the cursor fires mouseenter without the user moving the
  // mouse; ignore hover until there is a real move, so the first row stays
  // selected for the keyboard.
  const pointerMoved = useRef(false)

  // Focus goes back where it came from when the palette closes, and Tab stays
  // inside it while it is open — it is a modal, so it has to behave like one.
  const opener = useRef<HTMLElement | null>(null)
  const panelRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (open) {
      opener.current = document.activeElement as HTMLElement | null
      setQ(''); setActive(0); pointerMoved.current = false
    } else {
      opener.current?.focus?.()
    }
  }, [open])

  const available = (task: Task) => {
    if (!connected && !LOCAL_PAGES.includes(task.page)) return false
    if (hideUnavailable && task.page in access && access[task.page] === false) return false
    return true
  }

  // Match against the visible label plus the task's hidden keywords, so both
  // "add to channel" and "приватный канал" find the same task in either UI
  // language. Every typed word must hit somewhere; label hits rank higher.
  const hits = useMemo(() => {
    const query = q.trim().toLowerCase()
    const words = query.split(/\s+/).filter(Boolean)
    return TASKS.filter(available)
      .map((task) => {
        const label = t(`tasks.${task.id}`).toLowerCase()
        const page = t(`nav.${task.page}`).toLowerCase()
        const hay = `${label} ${page} ${task.keywords}`.toLowerCase()
        if (!words.every((w) => hay.includes(w))) return null
        const rank = !query ? 0 : label.startsWith(query) ? 2 : label.includes(query) ? 1 : 0
        return { task, rank }
      })
      .filter((x): x is { task: Task; rank: number } => x !== null)
      .sort((a, b) => b.rank - a.rank)
      .map((x) => x.task)
  }, [q, t, connected, readOnly, access, hideUnavailable])

  // Group headers only make sense while browsing the full list; once the
  // operator types, ranking beats grouping and the list stays flat.
  const grouped = useMemo(() => {
    if (q.trim()) return [{ group: null as string | null, items: hits }]
    return TASK_GROUPS.map((g) => ({ group: g as string | null, items: hits.filter((x) => x.group === g) }))
      .filter((s) => s.items.length > 0)
  }, [hits, q])

  // Keyboard navigation follows what is actually on screen, not the ranking
  // order — the two differ while the grouped (unfiltered) list is shown.
  const flat = useMemo(() => grouped.flatMap((s) => s.items), [grouped])

  useEffect(() => { setActive(0) }, [q])
  useEffect(() => {
    listRef.current?.querySelector('[data-active="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [active, q])

  if (!open) return null

  const pick = (task: Task) => {
    onClose()
    onNavigate(task.page)
    requestAction(task.action ?? null)
  }

  // Tab must not reach the page behind the overlay: the palette holds the only
  // two stops (the input and the active row), so the cycle stays inside.
  const trapTab = (e: React.KeyboardEvent) => {
    if (e.key !== 'Tab') return
    const focusable = panelRef.current?.querySelectorAll<HTMLElement>('input, button')
    if (!focusable || focusable.length === 0) return
    const first = focusable[0]
    const last = focusable[focusable.length - 1]
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus() }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus() }
  }

  const onKeyDown = (e: React.KeyboardEvent) => {
    trapTab(e)
    if (e.key === 'Escape') { e.preventDefault(); onClose(); return }
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive((i) => Math.min(i + 1, flat.length - 1)); return }
    if (e.key === 'ArrowUp') { e.preventDefault(); setActive((i) => Math.max(i - 1, 0)); return }
    if (e.key === 'Enter' && flat[active]) { e.preventDefault(); pick(flat[active]) }
  }

  let index = -1

  return (
    <div className="fixed inset-0 z-[200] flex items-start justify-center bg-black/50 p-4 pt-[12vh]" onClick={onClose}>
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-label={t('palette.open')}
        onKeyDown={trapTab}
        className="flex max-h-[70vh] w-full max-w-[640px] flex-col overflow-hidden rounded-2xl border border-[var(--border-strong)] bg-[var(--bg-elev)] shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center gap-2 border-b border-[var(--border)] px-4 py-3">
          <Search size={16} className="shrink-0 text-[var(--text-faint)]" />
          <input
            autoFocus
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onKeyDown={onKeyDown}
            placeholder={t('palette.placeholder')}
            className="w-full bg-transparent text-sm text-[var(--text)] placeholder:text-[var(--text-faint)] outline-none"
          />
          <kbd className="shrink-0 rounded border border-[var(--border)] px-1.5 py-0.5 text-[10px] text-[var(--text-faint)]">esc</kbd>
        </div>

        <div ref={listRef} className="min-h-0 flex-1 overflow-auto py-1" onMouseMove={() => { pointerMoved.current = true }}>
          {flat.length === 0 && (
            <div className="px-4 py-6 text-center text-sm text-[var(--text-faint)]">
              {connected ? t('palette.empty') : t('palette.connectFirst')}
            </div>
          )}
          {grouped.map((section) => (
            <div key={section.group ?? 'all'}>
              {section.group && (
                <div className="px-4 pb-1 pt-2 text-[11px] font-semibold uppercase tracking-wider text-[var(--text-faint)]">
                  {t(`palette.group.${section.group}`)}
                </div>
              )}
              {section.items.map((task) => {
                index += 1
                const i = index
                const isActive = i === active
                return (
                  <button
                    key={task.id}
                    data-active={isActive}
                    onMouseEnter={() => { if (pointerMoved.current) setActive(i) }}
                    onClick={() => pick(task)}
                    className={`flex w-full items-center gap-3 px-4 py-2 text-left ${isActive ? 'bg-[var(--accent)]/12' : ''}`}
                  >
                    <span className="min-w-0 flex-1 truncate text-sm text-[var(--text)]">{t(`tasks.${task.id}`)}</span>
                    {task.write && readOnly && (
                      <span className="flex shrink-0 items-center gap-1 text-[11px] text-[var(--warn)]">
                        <Lock size={11} /> {t('safety.readOnly')}
                      </span>
                    )}
                    <span className="shrink-0 text-xs text-[var(--text-faint)]">{t(`nav.${task.page}`)}</span>
                    {isActive && <CornerDownLeft size={13} className="shrink-0 text-[var(--text-faint)]" />}
                  </button>
                )
              })}
            </div>
          ))}
        </div>

        <div className="border-t border-[var(--border)] px-4 py-2 text-[11px] text-[var(--text-faint)]">
          {t('palette.hint')}
        </div>
      </div>
    </div>
  )
}
