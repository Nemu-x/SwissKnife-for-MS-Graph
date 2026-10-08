import { Component, type ReactNode } from 'react'
import i18n from '../i18n'

// A page that throws while rendering shows what broke instead of leaving the
// whole window empty; the menu stays usable. Keyed by page, so navigating
// away clears it.
export class PageErrorBoundary extends Component<{ children: ReactNode; onHome: () => void }, { error: Error | null }> {
  state = { error: null as Error | null }

  static getDerivedStateFromError(error: Error) {
    return { error }
  }

  componentDidCatch(error: Error) {
    console.error('page crashed', error)
  }

  render() {
    const { error } = this.state
    if (!error) return this.props.children
    return (
      <div className="flex h-full flex-col items-start gap-3 p-6">
        <h1 className="text-lg font-semibold">{i18n.t('crash.title')}</h1>
        <p className="text-sm text-[var(--text-dim)]">{i18n.t('crash.body')}</p>
        <pre className="max-w-full overflow-auto rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-3 text-xs text-[var(--danger)]">{String(error.message || error)}</pre>
        <div className="flex gap-2">
          <button className="rounded-lg bg-[var(--accent)] px-3 py-1.5 text-sm text-[var(--accent-fg)]" onClick={() => { this.setState({ error: null }); this.props.onHome() }}>
            {i18n.t('crash.home')}
          </button>
          <button className="rounded-lg border border-[var(--border)] px-3 py-1.5 text-sm" onClick={() => this.setState({ error: null })}>
            {i18n.t('crash.retry')}
          </button>
        </div>
      </div>
    )
  }
}
