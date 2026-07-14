// Async option loaders for EntityPicker — turn Graph collections into pickable
// options. Kept small; large tenants can still paste an ID manually.
import { api, type GraphObject } from './api'
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

export const loadSites = async (): Promise<Option[]> =>
  (await api.drive.sites('')).map((s: GraphObject) => ({
    value: s.id,
    label: s.displayName || s.name || s.id,
    sub: s.webUrl,
  }))
