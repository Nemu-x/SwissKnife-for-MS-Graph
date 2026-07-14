import { CheckCircle2, XCircle, Info, X } from 'lucide-react'
import { useStore } from '../lib/store'

export function Toasts() {
  const { toasts, dismiss } = useStore()
  return (
    <div className="pointer-events-none fixed bottom-4 right-4 z-50 flex flex-col gap-2">
      {toasts.map((tst) => {
        const Icon = tst.kind === 'ok' ? CheckCircle2 : tst.kind === 'err' ? XCircle : Info
        const color = tst.kind === 'ok' ? 'var(--ok)' : tst.kind === 'err' ? 'var(--danger)' : 'var(--accent)'
        return (
          <div
            key={tst.id}
            className="pointer-events-auto flex max-w-sm items-start gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] px-3 py-2 shadow-lg"
          >
            <Icon size={16} style={{ color }} className="mt-0.5 shrink-0" />
            <span className="text-sm text-[var(--text)]">{tst.text}</span>
            <button onClick={() => dismiss(tst.id)} className="ml-auto text-[var(--text-faint)] hover:text-[var(--text)]">
              <X size={14} />
            </button>
          </div>
        )
      })}
    </div>
  )
}
