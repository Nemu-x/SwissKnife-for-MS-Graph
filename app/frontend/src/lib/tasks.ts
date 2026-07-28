// Task registry for the command palette.
//
// The app is navigated by Graph domain (Users, Teams, Groups…), but operators
// arrive with a job to do ("add this person to a private channel"). A task
// names that job and points at the page — and, for ActionPage pages, the
// toolbar action — that performs it.
import type { PageId } from '../pages/registry'

export type TaskGroup = 'people' | 'access' | 'collab' | 'devices' | 'insight'

export interface Task {
  id: string // also the i18n key under `tasks.`
  group: TaskGroup
  page: PageId
  // ActionPage action id to trigger on arrival: a drawer action opens its
  // drawer, an immediate action runs. Pages without a toolbar just navigate.
  action?: string
  write?: boolean // needs write access — flagged while read-only mode is on
  // Extra match terms in both UI languages. Never rendered: the visible label
  // comes from i18n, this only widens what the operator can type to find it.
  keywords: string
}

export const TASK_GROUPS: TaskGroup[] = ['people', 'access', 'collab', 'devices', 'insight']

export const TASKS: Task[] = [
  // People
  { id: 'findUser', group: 'people', page: 'users', action: 'list', keywords: 'user search directory upn пользователь найти поиск список' },
  { id: 'userSnapshot', group: 'people', page: 'users', action: 'snapshot', keywords: 'user info about card пользователь карточка сведения' },
  { id: 'createUser', group: 'people', page: 'users', action: 'create', write: true, keywords: 'new account hire создать нового учётную запись' },
  { id: 'deleteUser', group: 'people', page: 'users', action: 'delete', write: true, keywords: 'remove account удалить учётную запись' },
  { id: 'blockUser', group: 'people', page: 'users', action: 'block', write: true, keywords: 'disable enable sign-in заблокировать разблокировать вход' },
  { id: 'resetPassword', group: 'people', page: 'users', action: 'password', write: true, keywords: 'password reset пароль сбросить сменить' },
  { id: 'revokeSessions', group: 'people', page: 'users', action: 'sessions', write: true, keywords: 'sessions sign out token выйти сессии токены' },
  { id: 'createTap', group: 'people', page: 'users', action: 'tap', write: true, keywords: 'tap temporary access pass onboarding временный пропуск код' },
  { id: 'resetMfa', group: 'people', page: 'users', action: 'mfa', write: true, keywords: 'mfa authenticator phone 2fa сбросить мфа телефон' },
  { id: 'setManager', group: 'people', page: 'users', action: 'manager', write: true, keywords: 'manager chief department location руководитель начальник отдел локация' },
  { id: 'patchUser', group: 'people', page: 'users', action: 'patch', write: true, keywords: 'job title department attributes json должность отдел поля атрибуты' },
  { id: 'mirrorAccess', group: 'people', page: 'users', action: 'copyAccess', write: true, keywords: 'same access as copy mirror clone like такой же доступ как у скопировать права выдать как' },
  { id: 'compareAccess', group: 'people', page: 'users', action: 'compareAccess', keywords: 'compare access difference diff сравнить доступы разница чем отличается' },
  { id: 'inviteGuest', group: 'people', page: 'users', action: 'invite', write: true, keywords: 'guest b2b invite external гость внешний пригласить' },
  { id: 'restoreUser', group: 'people', page: 'users', action: 'restore', write: true, keywords: 'restore deleted recycle восстановить удалённого корзина' },
  { id: 'onboard', group: 'people', page: 'playbooks', write: true, keywords: 'playbook hire new employee онбординг приём нового сотрудника плейбук' },
  { id: 'offboard', group: 'people', page: 'playbooks', write: true, keywords: 'playbook leaver fired quit оффбординг увольнение плейбук' },
  { id: 'bulkCsv', group: 'people', page: 'bulk', write: true, keywords: 'bulk csv mass import массово пачкой список импорт' },

  // Access
  { id: 'assignLicense', group: 'access', page: 'licensing', action: 'assign', write: true, keywords: 'license sku assign remove лицензия выдать снять' },
  { id: 'tenantLicenses', group: 'access', page: 'licensing', action: 'skus', keywords: 'sku stock seats free лицензии остаток места' },
  { id: 'userLicenses', group: 'access', page: 'licensing', action: 'userLicenses', keywords: 'license user which лицензии пользователя какие' },
  { id: 'addToGroup', group: 'access', page: 'groups', action: 'add', write: true, keywords: 'group member add owner группа добавить участник владелец' },
  { id: 'groupMembers', group: 'access', page: 'groups', action: 'members', keywords: 'group who members состав группы кто' },
  { id: 'createGroup', group: 'access', page: 'groups', action: 'create', write: true, keywords: 'group new m365 создать группу' },
  { id: 'adminRole', group: 'access', page: 'roles', action: 'grant', write: true, keywords: 'role admin grant revoke роль админ выдать забрать права' },
  { id: 'listRoles', group: 'access', page: 'roles', action: 'list', keywords: 'roles admins who список ролей админы' },
  { id: 'rotateSecret', group: 'access', page: 'apps', action: 'rotate', write: true, keywords: 'app secret rotate client приложение секрет ротация' },
  { id: 'expiringSecrets', group: 'access', page: 'apps', action: 'expiring', keywords: 'expiry certificate secret истекает сертификат секрет протухает' },

  // Collaboration
  { id: 'addToTeam', group: 'collab', page: 'teams', action: 'addToTeam', write: true, keywords: 'teams team member add команда добавить участник' },
  { id: 'addToChannel', group: 'collab', page: 'teams', action: 'addToChannel', write: true, keywords: 'private channel teams приватный канал добавить' },
  { id: 'teamMembership', group: 'collab', page: 'teams', action: 'whereIs', keywords: 'joined teams which в каких командах состоит' },
  { id: 'listChannels', group: 'collab', page: 'teams', action: 'teamContents', keywords: 'channels team list members каналы команды список состав' },
  { id: 'createChannel', group: 'collab', page: 'teams', action: 'createChannel', write: true, keywords: 'channel new private shared создать канал приватный' },
  { id: 'teamify', group: 'collab', page: 'teams', action: 'teamify', write: true, keywords: 'group to team teamify группу в команду' },
  { id: 'addToChat', group: 'collab', page: 'chats', action: 'members', write: true, keywords: 'chat member add чат добавить участник' },
  { id: 'readChat', group: 'collab', page: 'chats', action: 'read', keywords: 'chat messages read переписка сообщения чат прочитать' },
  { id: 'createChat', group: 'collab', page: 'chats', action: 'create', write: true, keywords: 'group chat new создать групповой чат' },
  { id: 'sendMail', group: 'collab', page: 'mail', action: 'send', write: true, keywords: 'mail send as отправить письмо от имени' },
  { id: 'readMailbox', group: 'collab', page: 'mail', action: 'read', keywords: 'mailbox inbox read почтовый ящик входящие посмотреть' },
  { id: 'createEvent', group: 'collab', page: 'mail', action: 'createEvent', write: true, keywords: 'calendar event meeting календарь встреча событие' },
  { id: 'browseDrive', group: 'collab', page: 'files', action: 'drive', keywords: 'onedrive sharepoint files site файлы диск сайт' },
  { id: 'uploadFile', group: 'collab', page: 'files', action: 'upload', write: true, keywords: 'upload file загрузить файл закинуть' },
  { id: 'fileActions', group: 'collab', page: 'files', action: 'item', write: true, keywords: 'download link share delete скачать ссылка расшарить удалить файл' },
  { id: 'backupOneDrive', group: 'collab', page: 'offboarding', write: true, keywords: 'backup copy onedrive leaver бэкап перенос файлов уволенного' },
  { id: 'freeSpace', group: 'collab', page: 'cleanup', write: true, keywords: 'duplicates versions storage quota дубликаты версии место квота' },

  // Devices
  { id: 'listDevices', group: 'devices', page: 'devices', action: 'list', keywords: 'entra devices list устройства список' },
  { id: 'deviceState', group: 'devices', page: 'devices', action: 'info', write: true, keywords: 'device enable disable устройство включить отключить' },
  { id: 'deleteDevice', group: 'devices', page: 'devices', action: 'delete', write: true, keywords: 'device delete remove entra удалить устройство' },
  { id: 'bitlocker', group: 'devices', page: 'devices', action: 'bitlocker', keywords: 'bitlocker recovery key ключ восстановления шифрование' },
  { id: 'intuneDevices', group: 'devices', page: 'intune', action: 'list', keywords: 'intune managed devices mdm устройства интюн' },
  { id: 'wipeDevice', group: 'devices', page: 'intune', action: 'wipe', write: true, keywords: 'wipe factory reset стереть сброс до заводских' },
  { id: 'retireDevice', group: 'devices', page: 'intune', action: 'retire', write: true, keywords: 'retire company data снять с учёта корпоративные данные' },
  { id: 'lockDevice', group: 'devices', page: 'intune', action: 'lock', write: true, keywords: 'lock stolen lost заблокировать украли потерял' },

  // Insight
  { id: 'signIns', group: 'insight', page: 'audit', action: 'signins', keywords: 'sign-in logs login failed логи входов вход не смог' },
  { id: 'directoryAudit', group: 'insight', page: 'audit', action: 'directory', keywords: 'audit who changed кто изменил аудит директории' },
  { id: 'usageReports', group: 'insight', page: 'reports', keywords: 'usage report activity csv отчёт использование активность' },
  { id: 'serviceHealth', group: 'insight', page: 'health', action: 'overview', keywords: 'outage incident status down лежит авария статус сервисов' },
  { id: 'messageCenter', group: 'insight', page: 'health', action: 'messages', keywords: 'message center announcements changes центр сообщений анонсы изменения' },
  { id: 'caPolicies', group: 'insight', page: 'security', action: 'ca', keywords: 'conditional access policy условный доступ политики' },
  { id: 'appConsents', group: 'insight', page: 'security', action: 'consents', keywords: 'consent oauth enterprise apps согласия разрешения приложений' },
  { id: 'runHistory', group: 'insight', page: 'history', keywords: 'history past runs resume история запусков продолжить' },
  { id: 'rawGraph', group: 'insight', page: 'raw', keywords: 'raw graph endpoint api request запрос вручную эндпоинт' },
  { id: 'dashboard', group: 'insight', page: 'dashboard', keywords: 'overview tenant counts обзор тенант дашборд' },
  { id: 'connectTenant', group: 'insight', page: 'connect', keywords: 'connect tenant profile switch подключиться тенант профиль сменить' },
  { id: 'settings', group: 'insight', page: 'settings', keywords: 'settings language theme accent настройки язык тема цвет' },
]
