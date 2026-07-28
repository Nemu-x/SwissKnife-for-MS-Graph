// Async option loaders for EntityPicker — turn Graph collections into pickable
// options. Kept small; large tenants can still paste an ID manually.
import { api, type GraphObject } from './api'
import { humanBytes } from './format'
import { skuFriendly } from './skuNames'
import type { Option } from '../components/MultiSelect'

export const loadUsers = async (): Promise<Option[]> =>
  (await api.users.list('', 200)).map((u: GraphObject) => ({
    value: u.userPrincipalName || u.id,
    label: u.displayName || u.userPrincipalName || u.id,
    sub: u.userPrincipalName || u.mail,
  }))

export const loadGroups = async (): Promise<Option[]> =>
  (await api.groups.list('', 0)).map((g: GraphObject) => ({
    value: g.id,
    label: g.displayName || g.id,
    sub: g.mail,
  }))

// Licenses by product name with seats used/total — nobody remembers SKU GUIDs.
export const loadSkus = async (): Promise<Option[]> =>
  (await api.licensing.skus()).map((s: GraphObject) => ({
    value: s.skuId,
    label: skuFriendly(s.skuPartNumber),
    sub: `${s.consumedUnits ?? '?'} / ${s.prepaidUnits?.enabled ?? '?'} · ${s.skuPartNumber}`,
  }))

export const loadTeams = async (): Promise<Option[]> =>
  (await api.teams.all()).map((t: GraphObject) => ({
    value: t.id,
    label: t.displayName || t.id,
    sub: t.description,
  }))

export const loadChannels = (teamId: string) => async (): Promise<Option[]> =>
  (await api.teams.channels(teamId)).map((c: GraphObject) => ({
    value: c.id,
    label: c.displayName || c.id,
    sub: c.membershipType,
  }))

// Only private and shared channels have their own membership — a standard
// channel inherits the team's members and Graph rejects adding anyone to it.
export const loadMembershipChannels = (teamId: string) => async (): Promise<Option[]> =>
  (await api.teams.channels(teamId))
    .filter((c: GraphObject) => c.membershipType && c.membershipType !== 'standard')
    .map((c: GraphObject) => ({
      value: c.id,
      label: c.displayName || c.id,
      sub: c.membershipType,
    }))

export const loadRoles = async (): Promise<Option[]> =>
  (await api.roles.list()).map((r: GraphObject) => ({
    value: r.id,
    label: r.displayName || r.id,
    sub: r.description,
  }))

export const loadDevices = async (): Promise<Option[]> =>
  (await api.devices.list('', 200)).map((d: GraphObject) => ({
    value: d.id,
    label: d.displayName || d.id,
    sub: [d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' '),
  }))

export const loadIntuneDevices = async (): Promise<Option[]> =>
  (await api.intune.devices(200)).map((d: GraphObject) => ({
    value: d.id,
    label: d.deviceName || d.id,
    sub: [d.operatingSystem, d.userPrincipalName].filter(Boolean).join(' · '),
  }))

export const loadChats = (user: string) => async (): Promise<Option[]> =>
  (await api.chats.listForPicker(user)).map((c) => ({
    value: c.id,
    label: c.label,
    sub: c.chatType,
  }))

export const loadSites = async (): Promise<Option[]> =>
  (await api.drive.sites(''))
    // This picker targets shared spaces; skip personal OneDrive sites.
    .filter((s: GraphObject) => !sitePath(s.webUrl).toLowerCase().includes('/personal/'))
    .map((s: GraphObject) => ({
      value: s.id,
      label: s.displayName || s.name || s.id,
      // show the site path (…/sites/Name) — it's what distinguishes same-named sites
      sub: sitePath(s.webUrl) || s.name,
    }))

// Like loadSites but prefixes each site with its storage used and sorts
// largest-first, so the operator can target the biggest sites. Slower: it reads
// each site's drive quota.
export const loadSitesBySize = async (): Promise<Option[]> =>
  (await api.drive.sitesWithUsage('')).map((s) => ({
    value: s.id,
    label: `${s.used >= 0 ? humanBytes(s.used) : '—'} · ${s.name}`,
    sub: sitePath(s.webUrl) || s.name,
  }))

function sitePath(webUrl?: string): string {
  if (!webUrl) return ''
  try {
    return new URL(webUrl).pathname || webUrl
  } catch {
    return webUrl
  }
}
