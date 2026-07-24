import { useState, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import {
  Plug, LayoutDashboard, PlayCircle, Users, KeyRound, ShieldCheck, Boxes, MessagesSquare, MessageCircle, Mail,
  FolderOpen, UserMinus, Smartphone, MonitorSmartphone, AppWindow, BarChart3, Sparkles, HeartPulse,
  ScrollText, TerminalSquare, Activity, Settings, Lock, Layers, ShieldAlert, History, ChevronDown,
} from 'lucide-react'
import { useStore } from '../lib/store'
import logo from '../assets/images/logo.png'
import type { PageId } from '../pages/registry'

type NavItem = { id: PageId; icon: ReactNode; key: string }

// Pages usable without a tenant connection.
const LOCAL_PAGES: PageId[] = ['connect', 'settings', 'history']

// Pinned entries above the groups.
const pinned: NavItem[] = [
  { id: 'connect', icon: <Plug size={17} />, key: 'nav.connect' },
  { id: 'dashboard', icon: <LayoutDashboard size={17} />, key: 'nav.dashboard' },
]

// Grouped navigation (nav-groups capability): collapsible sections, access
// hiding stays per item, a section with no visible items hides entirely.
const groups: { key: string; items: NavItem[] }[] = [
  {
    key: 'navGroups.identity',
    items: [
      { id: 'users', icon: <Users size={17} />, key: 'nav.users' },
      { id: 'licensing', icon: <KeyRound size={17} />, key: 'nav.licensing' },
      { id: 'roles', icon: <ShieldCheck size={17} />, key: 'nav.roles' },
      { id: 'groups', icon: <Boxes size={17} />, key: 'nav.groups' },
      { id: 'apps', icon: <AppWindow size={17} />, key: 'nav.apps' },
    ],
  },
  {
    key: 'navGroups.collab',
    items: [
      { id: 'teams', icon: <MessagesSquare size={17} />, key: 'nav.teams' },
      { id: 'chats', icon: <MessageCircle size={17} />, key: 'nav.chats' },
      { id: 'mail', icon: <Mail size={17} />, key: 'nav.mail' },
      { id: 'files', icon: <FolderOpen size={17} />, key: 'nav.files' },
    ],
  },
  {
    key: 'navGroups.devices',
    items: [
      { id: 'intune', icon: <Smartphone size={17} />, key: 'nav.intune' },
      { id: 'devices', icon: <MonitorSmartphone size={17} />, key: 'nav.devices' },
    ],
  },
  {
    key: 'navGroups.dataOps',
    items: [
      { id: 'playbooks', icon: <PlayCircle size={17} />, key: 'nav.playbooks' },
      { id: 'bulk', icon: <Layers size={17} />, key: 'nav.bulk' },
      { id: 'offboarding', icon: <UserMinus size={17} />, key: 'nav.offboarding' },
      { id: 'cleanup', icon: <Sparkles size={17} />, key: 'nav.cleanup' },
    ],
  },
  {
    key: 'navGroups.insights',
    items: [
      { id: 'security', icon: <ShieldAlert size={17} />, key: 'nav.security' },
      { id: 'reports', icon: <BarChart3 size={17} />, key: 'nav.reports' },
      { id: 'audit', icon: <ScrollText size={17} />, key: 'nav.audit' },
      { id: 'health', icon: <HeartPulse size={17} />, key: 'nav.health' },
      { id: 'activity', icon: <Activity size={17} />, key: 'nav.activity' },
      { id: 'history', icon: <History size={17} />, key: 'nav.history' },
    ],
  },
  {
    key: 'navGroups.system',
    items: [
      { id: 'raw', icon: <TerminalSquare size={17} />, key: 'nav.raw' },
      { id: 'settings', icon: <Settings size={17} />, key: 'nav.settings' },
    ],
  },
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
  const { connected, status, readOnly, access, hideUnavailable } = useStore()

  const [collapsed, setCollapsed] = useState<Record<string, boolean>>(() => {
    try { return JSON.parse(localStorage.getItem('navCollapsed') || '{}') } catch { return {} }
  })
  const toggleGroup = (key: string) => {
    setCollapsed((c) => {
      const next = { ...c, [key]: !c[key] }
      localStorage.setItem('navCollapsed', JSON.stringify(next))
      return next
    })
  }

  const itemVisible = (it: NavItem) => !hideUnavailable || !(it.id in access) || access[it.id] !== false

  const renderItem = (it: NavItem) => {
    const active = page === it.id
    const disabled = !LOCAL_PAGES.includes(it.id) && !connected
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
  }

  return (
    <div className="flex h-full">
      <aside className="flex w-56 shrink-0 flex-col border-r border-[var(--border)] bg-[var(--bg-elev)]">
        <div className="flex items-center gap-2.5 px-4 py-4">
          <img src={logo} alt="SwissKnife" className="h-7 w-7 rounded-md" />
          <span className="text-sm font-semibold leading-tight">SwissKnife<br /><span className="text-xs font-normal text-[var(--text-faint)]">for MS Graph</span></span>
        </div>
        <nav className="flex-1 overflow-y-auto px-2 py-1">
          {pinned.filter(itemVisible).map(renderItem)}
          {groups.map((g) => {
            const visible = g.items.filter(itemVisible)
            if (visible.length === 0) return null
            // The section holding the active page always shows its items.
            const isOpen = !collapsed[g.key] || visible.some((it) => it.id === page)
            return (
              <div key={g.key} className="mt-2">
                <button
                  onClick={() => toggleGroup(g.key)}
                  className="flex w-full items-center justify-between px-3 py-1 text-[11px] font-semibold uppercase tracking-wider text-[var(--text-faint)] hover:text-[var(--text-dim)]"
                >
                  <span className="truncate">{t(g.key)}</span>
                  <ChevronDown size={12} className={`shrink-0 transition-transform ${isOpen ? '' : '-rotate-90'}`} />
                </button>
                {isOpen && visible.map(renderItem)}
              </div>
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
