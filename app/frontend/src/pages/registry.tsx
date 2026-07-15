import type { ComponentType } from 'react'
import { ConnectPage } from './ConnectPage'
import { DashboardPage } from './DashboardPage'
import { PlaybooksPage } from './PlaybooksPage'
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
import { DevicesPage } from './DevicesPage'
import { AppsPage } from './AppsPage'
import { ReportsPage } from './ReportsPage'
import { BulkPage } from './BulkPage'
import { SecurityPage } from './SecurityPage'
import { CleanupPage } from './CleanupPage'
import { HealthPage } from './HealthPage'
import { AuditPage } from './AuditPage'
import { RawPage } from './RawPage'
import { ActivityPage } from './ActivityPage'
import { SettingsPage } from './SettingsPage'

export type PageId =
  | 'connect' | 'dashboard' | 'playbooks' | 'bulk' | 'users' | 'licensing' | 'roles' | 'groups' | 'teams' | 'chats'
  | 'mail' | 'files' | 'offboarding' | 'intune' | 'devices' | 'apps' | 'security' | 'reports' | 'cleanup' | 'health'
  | 'audit' | 'raw' | 'activity' | 'settings'

export const pages: Record<PageId, ComponentType> = {
  connect: ConnectPage,
  dashboard: DashboardPage,
  playbooks: PlaybooksPage,
  bulk: BulkPage,
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
  devices: DevicesPage,
  apps: AppsPage,
  security: SecurityPage,
  reports: ReportsPage,
  cleanup: CleanupPage,
  health: HealthPage,
  audit: AuditPage,
  raw: RawPage,
  activity: ActivityPage,
  settings: SettingsPage,
}
