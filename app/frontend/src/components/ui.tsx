import type { ButtonHTMLAttributes, InputHTMLAttributes, ReactNode, SelectHTMLAttributes, TextareaHTMLAttributes } from 'react'

type Variant = 'primary' | 'ghost' | 'danger' | 'subtle'

const variants: Record<Variant, string> = {
  primary: 'bg-[var(--accent)] text-[var(--accent-fg)] hover:bg-[var(--accent-hover)] border-transparent',
  danger: 'bg-[var(--danger)] text-white hover:bg-[var(--danger-hover)] border-transparent',
  ghost: 'bg-transparent text-[var(--text)] hover:bg-[var(--bg-elev-2)] border-[var(--border)]',
  subtle: 'bg-[var(--bg-elev-2)] text-[var(--text)] hover:bg-[var(--border)] border-transparent',
}

export function Button({
  variant = 'subtle',
  className = '',
  children,
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { variant?: Variant }) {
  return (
    <button
      className={`inline-flex items-center justify-center gap-2 rounded-lg border px-3 py-1.5 text-sm font-medium
        transition-colors disabled:cursor-not-allowed disabled:opacity-50 ${variants[variant]} ${className}`}
      {...props}
    >
      {children}
    </button>
  )
}

export function Input({ className = '', ...props }: InputHTMLAttributes<HTMLInputElement>) {
  return (
    <input
      className={`w-full rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm
        text-[var(--text)] placeholder:text-[var(--text-faint)] focus:border-[var(--accent)] outline-none ${className}`}
      {...props}
    />
  )
}

export function Textarea({ className = '', ...props }: TextareaHTMLAttributes<HTMLTextAreaElement>) {
  return (
    <textarea
      className={`w-full rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm font-mono
        text-[var(--text)] placeholder:text-[var(--text-faint)] focus:border-[var(--accent)] outline-none ${className}`}
      {...props}
    />
  )
}

export function Select({ className = '', children, ...props }: SelectHTMLAttributes<HTMLSelectElement>) {
  return (
    <select
      className={`rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm
        text-[var(--text)] focus:border-[var(--accent)] outline-none ${className}`}
      {...props}
    >
      {children}
    </select>
  )
}

export function Field({ label, hint, children }: { label: string; hint?: string; children: ReactNode }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs font-medium text-[var(--text-dim)]">{label}</span>
      {children}
      {hint && <span className="text-xs text-[var(--text-faint)]">{hint}</span>}
    </label>
  )
}

export function Card({ title, actions, children, className = '' }: {
  title?: ReactNode
  actions?: ReactNode
  children: ReactNode
  className?: string
}) {
  return (
    <section className={`rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] ${className}`}>
      {(title || actions) && (
        <header className="flex items-center justify-between gap-2 border-b border-[var(--border)] px-4 py-3">
          <h2 className="text-sm font-semibold">{title}</h2>
          <div className="flex items-center gap-2">{actions}</div>
        </header>
      )}
      <div className="p-4">{children}</div>
    </section>
  )
}

export function Badge({ kind = 'neutral', children }: { kind?: 'ok' | 'warn' | 'danger' | 'neutral'; children: ReactNode }) {
  const colors = {
    ok: 'bg-[var(--ok)]/15 text-[var(--ok)]',
    warn: 'bg-[var(--warn)]/15 text-[var(--warn)]',
    danger: 'bg-[var(--danger)]/15 text-[var(--danger)]',
    neutral: 'bg-[var(--bg-elev-2)] text-[var(--text-dim)]',
  }[kind]
  return <span className={`rounded-md px-2 py-0.5 text-xs font-medium ${colors}`}>{children}</span>
}

export function Spinner() {
  return (
    <span className="inline-block h-4 w-4 animate-spin rounded-full border-2 border-[var(--text-faint)] border-t-[var(--accent)]" />
  )
}

export function ErrorNote({ children }: { children: ReactNode }) {
  return (
    <div className="rounded-lg border border-[var(--danger)]/40 bg-[var(--danger)]/10 px-3 py-2 text-sm text-[var(--danger)]">
      {children}
    </div>
  )
}
