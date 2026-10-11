import type { PageId } from '../pages/registry'

// What the operator manages: Microsoft 365 (the tenant), on-premises Active
// Directory, or both. A part that is switched off disappears from the menu
// and the task palette; nothing about it is deleted.
export type Workspaces = { cloud: boolean; onprem: boolean }

export const DEFAULT_WORKSPACES: Workspaces = { cloud: true, onprem: false }

// Pages that belong to neither part.
const SHARED: PageId[] = ['settings', 'history']
const ONPREM: PageId[] = ['onprem']

export function pageEnabled(p: PageId, ws: Workspaces): boolean {
  if (SHARED.includes(p)) return true
  if (ONPREM.includes(p)) return ws.onprem
  return ws.cloud
}

// Where the app starts (and falls back to) for the parts switched on.
export const homePage = (ws: Workspaces): PageId => (ws.cloud ? 'connect' : 'onprem')

export function loadWorkspaces(): Workspaces {
  try {
    const v = JSON.parse(localStorage.getItem('workspaces') || 'null')
    if (v && typeof v.cloud === 'boolean' && typeof v.onprem === 'boolean' && (v.cloud || v.onprem)) return v
  } catch { /* first run or blocked storage */ }
  return DEFAULT_WORKSPACES
}

export function saveWorkspaces(ws: Workspaces) {
  try { localStorage.setItem('workspaces', JSON.stringify(ws)) } catch { /* not persisted */ }
}
