import { useState, type ReactNode } from 'react'
import { BrowserOpenURL } from '../../wailsjs/runtime/runtime'
import { loadCompactNav, saveCompactNav } from '../lib/navprefs'
import { TRIBUTE_URL } from '../lib/support'
import { useTranslation } from 'react-i18next'
import {
  Plug, LayoutDashboard, PlayCircle, Users, KeyRound, ShieldCheck, Boxes, MessagesSquare, MessageCircle, Mail,
  FolderOpen, UserMinus, Smartphone, MonitorSmartphone, AppWindow, BarChart3, Sparkles, HeartPulse,
  ScrollText, TerminalSquare, Settings, Lock, Layers, ShieldAlert, History, ChevronDown, Search, ShieldHalf, Compass, Server,
  PanelLeftClose, PanelLeftOpen, Heart,
} from 'lucide-react'
import { useStore } from '../lib/store'
import { pageEnabled } from '../lib/workspaces'
import logo from '../assets/images/logo.png'
import type { PageId } from '../pages/registry'

export type NavItem = { id: PageId; icon: ReactNode; key: string }

// Pages usable without a tenant connection — the single source of truth for
// both the nav enablement here and the redirect guard in App.tsx.
export const LOCAL_PAGES: PageId[] = ['connect', 'settings', 'history', 'onprem']

// Pinned entries above the groups.
export const NAV_PINNED: NavItem[] = [
  { id: 'connect', icon: <Plug size={17} />, key: 'nav.connect' },
  { id: 'dashboard', icon: <LayoutDashboard size={17} />, key: 'nav.dashboard' },
  { id: 'explorer', icon: <Compass size={17} />, key: 'nav.explorer' },
]

// Grouped navigation (nav-groups capability): collapsible sections, access
// hiding stays per item, a section with no visible items hides entirely.
export const NAV_GROUPS: { key: string; items: NavItem[] }[] = [
  {
    key: 'navGroups.identity',
    items: [
      { id: 'users', icon: <Users size={17} />, key: 'nav.users' },
      { id: 'licensing', icon: <KeyRound size={17} />, key: 'nav.licensing' },
      { id: 'roles', icon: <ShieldCheck size={17} />, key: 'nav.roles' },
      { id: 'groups', icon: <Boxes size={17} />, key: 'nav.groups' },
      { id: 'apps', icon: <AppWindow size={17} />, key: 'nav.apps' },
      { id: 'onprem', icon: <Server size={17} />, key: 'nav.onprem' },
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
      // The local action log lives inside Run history now — one place for
      // "what did this app do", instead of two journals in two tabs.
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
  onOpenPalette,
  children,
}: {
  page: PageId
  onNavigate: (p: PageId) => void
  onOpenPalette?: () => void
  children: ReactNode
}) {
  const { t } = useTranslation()
  const { connected, status, readOnly, access, hideUnavailable, workspaces, hiddenPages } = useStore()
  // Icons only, names on hover: more room for the page.
  const [compact, setCompact] = useState(loadCompactNav)
  const toggleCompact = () => { setTip(null); setCompact((c) => { saveCompactNav(!c); return !c }) }
  // The hover label is drawn outside the scrolling nav, so it is never clipped.
  const [tip, setTip] = useState<{ text: string; top: number } | null>(null)
  const tipProps = (text: string) => compact ? {
    onMouseEnter: (e: React.MouseEvent<HTMLElement>) => { const r = e.currentTarget.getBoundingClientRect(); setTip({ text, top: r.top + r.height / 2 }) },
    onMouseLeave: () => setTip(null),
    onFocus: (e: React.FocusEvent<HTMLElement>) => { const r = e.currentTarget.getBoundingClientRect(); setTip({ text, top: r.top + r.height / 2 }) },
    onBlur: () => setTip(null),
    'aria-label': text,
  } : {}

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

  // A hidden page still shows while it is open, so the menu never loses the
  // operator's place.
  const itemVisible = (it: NavItem) => pageEnabled(it.id, workspaces) && (!hiddenPages.includes(it.id) || it.id === page) &&
    (!hideUnavailable || !(it.id in access) || access[it.id] !== false)

  const renderItem = (it: NavItem) => {
    const active = page === it.id
    const disabled = !LOCAL_PAGES.includes(it.id) && !connected
    return (
      <button
        key={it.id}
        disabled={disabled}
        onClick={() => onNavigate(it.id)}
        {...tipProps(t(it.key))}
        className={`mb-0.5 flex w-full items-center rounded-lg py-2 text-sm transition-colors
          ${compact ? 'justify-center px-0' : 'gap-2.5 px-3'}
          ${active ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--text)]'}
          disabled:cursor-not-allowed disabled:opacity-35`}
      >
        {it.icon}
        {!compact && <span className="truncate">{t(it.key)}</span>}
      </button>
    )
  }

  return (
    <div className="flex h-full">
      <aside className={`flex shrink-0 flex-col border-r border-[var(--border)] bg-[var(--bg-elev)] transition-[width] ${compact ? 'w-14' : 'w-56'}`}>
        <div className={`flex items-center gap-2.5 py-4 ${compact ? 'justify-center px-0' : 'px-4'}`}>
          <img src={logo} alt="SwissKnife" className="h-7 w-7 rounded-md" />
          {!compact && <span className="text-sm font-semibold leading-tight">SwissKnife<br /><span className="text-xs font-normal text-[var(--text-faint)]">for MS Graph</span></span>}
        </div>
        {onOpenPalette && compact && (
          <div className="px-2 pb-2">
            <button onClick={onOpenPalette} {...tipProps(t('palette.open') + ' (Ctrl+K)')}
              className="flex w-full justify-center rounded-lg border border-[var(--border)] bg-[var(--bg)] py-2 text-[var(--text-faint)] hover:border-[var(--accent)] hover:text-[var(--text-dim)]">
              <Search size={16} />
            </button>
          </div>
        )}
        {onOpenPalette && !compact && (
          <div className="px-2 pb-2">
            <button
              onClick={onOpenPalette}
              className="flex w-full items-start gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm text-[var(--text-faint)] hover:border-[var(--accent)] hover:text-[var(--text-dim)]"
            >
              <Search size={15} className="mt-0.5 shrink-0" />
              <span className="min-w-0 flex-1 text-left leading-snug">{t('palette.open')}</span>
              <kbd className="mt-0.5 shrink-0 rounded border border-[var(--border)] px-1 py-0.5 text-[10px] leading-none">Ctrl K</kbd>
            </button>
          </div>
        )}
        <nav className="flex-1 overflow-y-auto px-2 py-1">
          {NAV_PINNED.filter(itemVisible).map(renderItem)}
          {NAV_GROUPS.map((g) => {
            const visible = g.items.filter(itemVisible)
            if (visible.length === 0) return null
            // Icons only: sections become thin separators, every item shows.
            if (compact) {
              return <div key={g.key} className="mt-2 border-t border-[var(--border)] pt-2">{visible.map(renderItem)}</div>
            }
            // The section holding the active page always shows its items.
            const isOpen = !collapsed[g.key] || visible.some((it) => it.id === page)
            return (
              <div key={g.key} className="mt-2">
                <button
                  onClick={() => toggleGroup(g.key)}
                  aria-expanded={isOpen}
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
        <div className={`border-t border-[var(--border)] py-3 text-xs ${compact ? 'flex flex-col items-center gap-2 px-0' : 'px-4'}`}>
          {compact ? (
            <>
              {workspaces.cloud && (
                <span {...tipProps(connected ? status?.profileName || '' : t('common.notConnected'))} tabIndex={0}
                  className={`h-2.5 w-2.5 rounded-full ${connected ? 'bg-[var(--ok)]' : 'bg-[var(--text-faint)]'}`} />
              )}
              {readOnly && <span {...tipProps(t('safety.readOnly'))} tabIndex={0} className="text-[var(--warn)]"><Lock size={13} /></span>}
              {connected && (status as any)?.policy && ((status as any).policy.maxDanger || (status as any).policy.allowedGroups?.length > 0) && (
                <span {...tipProps(t('connect.limits.badge'))} tabIndex={0} className="text-[var(--warn)]"><ShieldHalf size={13} /></span>
              )}
            </>
          ) : <>
          {workspaces.cloud && <div className="flex items-center gap-1.5">
            <span className={`h-2 w-2 rounded-full ${connected ? 'bg-[var(--ok)]' : 'bg-[var(--text-faint)]'}`} />
            <span className="truncate text-[var(--text-dim)]">
              {connected ? status?.profileName : t('common.notConnected')}
            </span>
          </div>}
          {readOnly && (
            <div className="mt-1.5 flex items-center gap-1 text-[var(--warn)]">
              <Lock size={12} /> {t('safety.readOnly')}
            </div>
          )}
          {connected && (status as any)?.policy && ((status as any).policy.maxDanger || (status as any).policy.allowedGroups?.length > 0) && (
            <div className="mt-1.5 flex items-center gap-1 text-[var(--warn)]" title={t('connect.limits.active')}>
              <ShieldHalf size={12} /> {t('connect.limits.badge')}
            </div>
          )}
          </>}
          <div className={`flex items-center ${compact ? 'flex-col gap-1' : 'mt-2 justify-between'}`}>
            <button onClick={() => BrowserOpenURL(TRIBUTE_URL)} {...tipProps(t('support.short'))}
              className="flex items-center gap-1.5 rounded-md px-1.5 py-1 text-[var(--text-faint)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--danger)]">
              <Heart size={14} />{!compact && <span>{t('support.short')}</span>}
            </button>
            <button onClick={toggleCompact} {...tipProps(t(compact ? 'nav.expand' : 'nav.collapse'))} title={compact ? undefined : t('nav.collapse')}
              className="rounded-md p-1 text-[var(--text-faint)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--text)]">
              {compact ? <PanelLeftOpen size={15} /> : <PanelLeftClose size={15} />}
            </button>
          </div>
        </div>
      </aside>
      {tip && (
        <div role="tooltip" style={{ top: tip.top, left: 60 }}
          className="pointer-events-none fixed z-50 -translate-y-1/2 whitespace-nowrap rounded-md border border-[var(--border)] bg-[var(--bg-elev-2)] px-2 py-1 text-xs text-[var(--text)] shadow-lg">
          {tip.text}
        </div>
      )}
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
