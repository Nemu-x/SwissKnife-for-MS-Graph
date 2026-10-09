import type { PageId } from '../pages/registry'

// Menu preferences on this device: pages hidden from the sidebar (still
// reachable through Ctrl+K) and the icons-only sidebar.

// Pages that can never be hidden: the way back to every other setting.
export const ALWAYS_SHOWN: PageId[] = ['settings']

// "Essentials": what most admins use every day; the rest stays one Ctrl+K away.
export const ESSENTIAL_PAGES: PageId[] = [
  'connect', 'dashboard', 'explorer', 'users', 'groups', 'licensing', 'onprem',
  'teams', 'mail', 'intune', 'playbooks', 'offboarding', 'security', 'reports', 'history', 'settings',
]

function read<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(key)
    return v === null ? fallback : (JSON.parse(v) as T)
  } catch { return fallback }
}

function write(key: string, v: unknown) {
  try { localStorage.setItem(key, JSON.stringify(v)) } catch { /* not persisted */ }
}

export const loadHiddenPages = (): PageId[] => {
  const v = read<unknown>('navHidden', [])
  return Array.isArray(v) ? (v as PageId[]).filter((p) => !ALWAYS_SHOWN.includes(p)) : []
}
export const saveHiddenPages = (v: PageId[]) => write('navHidden', v.filter((p) => !ALWAYS_SHOWN.includes(p)))

export const loadCompactNav = (): boolean => read<boolean>('navCompact', false) === true
export const saveCompactNav = (v: boolean) => write('navCompact', v)
