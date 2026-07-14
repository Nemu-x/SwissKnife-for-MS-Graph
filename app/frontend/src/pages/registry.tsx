import type { ComponentType } from 'react'
import { ConnectPage } from './ConnectPage'
import { UsersPage } from './UsersPage'
import { LicensingPage } from './LicensingPage'
import { GroupsPage } from './GroupsPage'
import { TeamsPage } from './TeamsPage'
import { ChatsPage } from './ChatsPage'
import { MailPage } from './MailPage'
import { FilesPage } from './FilesPage'
import { IntunePage } from './IntunePage'
import { AuditPage } from './AuditPage'
import { RawPage } from './RawPage'
import { ActivityPage } from './ActivityPage'
import { SettingsPage } from './SettingsPage'

export type PageId =
  | 'connect' | 'users' | 'licensing' | 'groups' | 'teams' | 'chats'
  | 'mail' | 'files' | 'intune' | 'audit' | 'raw' | 'activity' | 'settings'

export const pages: Record<PageId, ComponentType> = {
  connect: ConnectPage,
  users: UsersPage,
  licensing: LicensingPage,
  groups: GroupsPage,
  teams: TeamsPage,
  chats: ChatsPage,
  mail: MailPage,
  files: FilesPage,
  intune: IntunePage,
  audit: AuditPage,
  raw: RawPage,
  activity: ActivityPage,
  settings: SettingsPage,
}
