import { useEffect, useRef } from 'react'
import { useTranslation } from 'react-i18next'
import { Loader2, X, Eraser } from 'lucide-react'
import type { JobState } from '../lib/store'

// Live console for a long-running job (offboarding copy, cleanup scan). It reads
// the job straight from the store, so it stays populated across page navigation
// and offers a Cancel button while the job is running.
export function JobConsole({
  job,
  onCancel,
  onClear,
}: {
  job: JobState
  onCancel: () => void
  onClear: () => void
}) {
  const { t } = useTranslation()
  const boxRef = useRef<HTMLDivElement>(null)

  // Keep the newest log line in view as it streams.
  useEffect(() => {
    const el = boxRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [job.log.length])

  if (job.log.length === 0 && !job.running) return null

  const color = (l: string) =>
    l.includes('✗') ? 'text-[var(--danger)]' : l.includes('✓') ? 'text-[var(--ok)]' : l.includes('⏹') ? 'text-[var(--warn)]' : ''

  return (
    <div className="rounded-lg border border-[var(--border)] bg-[var(--bg)]">
      <div className="flex items-center justify-between gap-2 border-b border-[var(--border)] px-3 py-1.5">
        <span className="flex items-center gap-2 text-xs font-semibold text-[var(--text-dim)]">
          {t('common.log')}
          {job.running && (
            <span className="flex items-center gap-1.5 font-normal text-[var(--accent2)]">
              <Loader2 size={12} className="animate-spin" /> {job.progress || t('common.working')}
            </span>
          )}
        </span>
        {job.running ? (
          <button onClick={onCancel} disabled={job.canceled}
            className="flex items-center gap-1 rounded-md border border-[var(--danger)]/40 px-2 py-0.5 text-xs text-[var(--danger)] hover:bg-[var(--danger)]/10 disabled:opacity-50">
            <X size={12} /> {job.canceled ? t('common.canceling') : t('common.cancel')}
          </button>
        ) : (
          <button onClick={onClear}
            className="flex items-center gap-1 rounded-md px-2 py-0.5 text-xs text-[var(--text-faint)] hover:text-[var(--text-dim)]">
            <Eraser size={12} /> {t('common.clear')}
          </button>
        )}
      </div>
      <div ref={boxRef} className="max-h-56 overflow-auto p-2 font-mono text-xs leading-relaxed text-[var(--text-dim)]">
        {job.log.map((l, i) => (
          <div key={i} className={color(l)}>{l}</div>
        ))}
        {job.running && job.progress && <div className="text-[var(--text-faint)]">… {job.progress}</div>}
      </div>
    </div>
  )
}
