import type { ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import {
  Plug, Users, KeyRound, Boxes, MessagesSquare, MessageCircle, Mail,
  FolderOpen, Smartphone, ScrollText, TerminalSquare, Activity, Settings, Lock,
} from 'lucide-react'
import { useStore } from '../lib/store'
import type { PageId } from '../pages/registry'

const items: { id: PageId; icon: ReactNode; key: string }[] = [
  { id: 'connect', icon: <Plug size={17} />, key: 'nav.connect' },
  { id: 'users', icon: <Users size={17} />, key: 'nav.users' },
  { id: 'licensing', icon: <KeyRound size={17} />, key: 'nav.licensing' },
  { id: 'groups', icon: <Boxes size={17} />, key: 'nav.groups' },
  { id: 'teams', icon: <MessagesSquare size={17} />, key: 'nav.teams' },
  { id: 'chats', icon: <MessageCircle size={17} />, key: 'nav.chats' },
  { id: 'mail', icon: <Mail size={17} />, key: 'nav.mail' },
  { id: 'files', icon: <FolderOpen size={17} />, key: 'nav.files' },
  { id: 'intune', icon: <Smartphone size={17} />, key: 'nav.intune' },
  { id: 'audit', icon: <ScrollText size={17} />, key: 'nav.audit' },
  { id: 'raw', icon: <TerminalSquare size={17} />, key: 'nav.raw' },
  { id: 'activity', icon: <Activity size={17} />, key: 'nav.activity' },
  { id: 'settings', icon: <Settings size={17} />, key: 'nav.settings' },
]

export function Layout({
  page,
  onNavigate,
  children,
}: {
  page: PageId
  onNavigate: (p: PageId) => void
  children: ReactNode
}) {
  const { t } = useTranslation()
  const { connected, status, readOnly } = useStore()

  return (
    <div className="flex h-full">
      <aside className="flex w-56 shrink-0 flex-col border-r border-[var(--border)] bg-[var(--bg-elev)]">
        <div className="flex items-center gap-2 px-4 py-4">
          <span className="text-lg">🗡️</span>
          <span className="text-sm font-semibold leading-tight">SwissKnife<br /><span className="text-xs font-normal text-[var(--text-faint)]">for MS Graph</span></span>
        </div>
        <nav className="flex-1 overflow-y-auto px-2 py-1">
          {items.map((it) => {
            const active = page === it.id
            const disabled = it.id !== 'connect' && it.id !== 'settings' && !connected
            return (
              <button
                key={it.id}
                disabled={disabled}
                onClick={() => onNavigate(it.id)}
                className={`mb-0.5 flex w-full items-center gap-2.5 rounded-lg px-3 py-2 text-sm transition-colors
                  ${active ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--text)]'}
                  disabled:cursor-not-allowed disabled:opacity-35`}
              >
                {it.icon}
                <span className="truncate">{t(it.key)}</span>
              </button>
            )
          })}
        </nav>
        <div className="border-t border-[var(--border)] px-4 py-3 text-xs">
          <div className="flex items-center gap-1.5">
            <span className={`h-2 w-2 rounded-full ${connected ? 'bg-[var(--ok)]' : 'bg-[var(--text-faint)]'}`} />
            <span className="truncate text-[var(--text-dim)]">
              {connected ? status?.profileName : t('common.notConnected')}
            </span>
          </div>
          {readOnly && (
            <div className="mt-1.5 flex items-center gap-1 text-[var(--warn)]">
              <Lock size={12} /> {t('safety.readOnly')}
            </div>
          )}
        </div>
      </aside>
      <main className="min-w-0 flex-1 overflow-hidden bg-[var(--bg)]">{children}</main>
    </div>
  )
}

export function Page({ title, subtitle, children }: { title: string; subtitle?: string; children: ReactNode }) {
  return (
    <div className="flex h-full flex-col">
      <header className="shrink-0 border-b border-[var(--border)] px-6 py-4">
        <h1 className="text-lg font-semibold">{title}</h1>
        {subtitle && <p className="mt-0.5 text-sm text-[var(--text-dim)]">{subtitle}</p>}
      </header>
      <div className="min-h-0 flex-1 overflow-auto p-6">{children}</div>
    </div>
  )
}
