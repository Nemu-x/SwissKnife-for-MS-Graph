import type { ReactNode } from 'react'

// Left pane is the controls, right pane is the ResultView.
export function TwoPane({ controls, result }: { controls: ReactNode; result: ReactNode }) {
  return (
    <div className="grid h-full grid-cols-1 gap-4 lg:grid-cols-[minmax(320px,380px)_1fr]">
      <div className="flex flex-col gap-4 overflow-auto">{controls}</div>
      <div className="min-h-[300px] overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--bg-elev)]">
        {result}
      </div>
    </div>
  )
}
