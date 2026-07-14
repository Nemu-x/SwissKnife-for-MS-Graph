import type { ComponentType } from 'react'
import { ConnectPage } from './ConnectPage'
import { DashboardPage } from './DashboardPage'
import { UsersPage } from './UsersPage'
import { LicensingPage } from './LicensingPage'
import { RolesPage } from './RolesPage'
import { GroupsPage } from './GroupsPage'
import { TeamsPage } from './TeamsPage'
import { ChatsPage } from './ChatsPage'
import { MailPage } from './MailPage'
import { FilesPage } from './FilesPage'
import { OffboardingPage } from './OffboardingPage'
import { IntunePage } from './IntunePage'
import { AuditPage } from './AuditPage'
import { RawPage } from './RawPage'
import { ActivityPage } from './ActivityPage'
import { SettingsPage } from './SettingsPage'

export type PageId =
  | 'connect' | 'dashboard' | 'users' | 'licensing' | 'roles' | 'groups' | 'teams' | 'chats'
  | 'mail' | 'files' | 'offboarding' | 'intune' | 'audit' | 'raw' | 'activity' | 'settings'

export const pages: Record<PageId, ComponentType> = {
  connect: ConnectPage,
  dashboard: DashboardPage,
  users: UsersPage,
  licensing: LicensingPage,
  roles: RolesPage,
  groups: GroupsPage,
  teams: TeamsPage,
  chats: ChatsPage,
  mail: MailPage,
  files: FilesPage,
  offboarding: OffboardingPage,
  intune: IntunePage,
  audit: AuditPage,
  raw: RawPage,
  activity: ActivityPage,
  settings: SettingsPage,
}
