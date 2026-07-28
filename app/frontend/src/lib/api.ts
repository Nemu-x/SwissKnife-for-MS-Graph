// Typed wrapper over the generated Wails bindings.
// json.RawMessage arrives from Go as real JSON (mistyped as number[] in .d.ts),
// so we cast it to proper types here.
import * as Connect from '../../wailsjs/go/services/ConnectService'
import * as Dashboard from '../../wailsjs/go/services/DashboardService'
import * as Playbook from '../../wailsjs/go/services/PlaybookService'
import * as Mirror from '../../wailsjs/go/services/MirrorService'
import * as Access from '../../wailsjs/go/services/AccessService'
import * as Users from '../../wailsjs/go/services/UsersService'
import * as AuthMethods from '../../wailsjs/go/services/AuthMethodsService'
import * as Roles from '../../wailsjs/go/services/RolesService'
import * as Licensing from '../../wailsjs/go/services/LicensingService'
import * as Groups from '../../wailsjs/go/services/GroupsService'
import * as Teams from '../../wailsjs/go/services/TeamsService'
import * as Chats from '../../wailsjs/go/services/ChatsService'
import * as Mail from '../../wailsjs/go/services/MailService'
import * as Calendar from '../../wailsjs/go/services/CalendarService'
import * as Drive from '../../wailsjs/go/services/DriveService'
import * as Intune from '../../wailsjs/go/services/IntuneService'
import * as Audit from '../../wailsjs/go/services/AuditService'
import * as Raw from '../../wailsjs/go/services/RawService'
import * as Devices from '../../wailsjs/go/services/DevicesService'
import * as Apps from '../../wailsjs/go/services/AppsService'
import * as Reports from '../../wailsjs/go/services/ReportsService'
import * as Cleanup from '../../wailsjs/go/services/CleanupService'
import * as ServiceHealth from '../../wailsjs/go/services/ServiceHealthService'
import * as Security from '../../wailsjs/go/services/SecurityService'
import * as Update from '../../wailsjs/go/services/UpdateService'
import * as Journal from '../../wailsjs/go/services/JournalService'
import * as Notify from '../../wailsjs/go/services/NotifyService'
import type { journal, secrets, services } from '../../wailsjs/go/models'
import { parseErr, type ParsedError } from './graphError'
import i18n from '../i18n'

export type GraphObject = Record<string, any>
export type Profile = secrets.Profile
export type Status = services.Status
export type ConnectRequest = services.ConnectRequest

// A Go service that returns a nil slice marshals to `null`, not `[]` — coerce
// here so no caller has to guard a .map() against an empty collection.
const list = (p: Promise<any>) => p.then((v) => (v ?? []) as GraphObject[])
const one = (p: Promise<any>) => p as Promise<GraphObject>

export const api = {
  connect: {
    profiles: () => Connect.Profiles() as Promise<Profile[]>,
    saveProfile: (p: Partial<Profile>, secret: string) =>
      Connect.SaveProfile(p as Profile, secret) as Promise<Profile>,
    deleteProfile: (id: string) => Connect.DeleteProfile(id),
    connect: (req: Partial<ConnectRequest>) => Connect.Connect(req as ConnectRequest) as Promise<Status>,
    disconnect: () => Connect.Disconnect(),
    status: () => Connect.GetStatus() as Promise<Status>,
    setReadOnly: (v: boolean) => Connect.SetReadOnly(v) as Promise<Status>,
    domains: () => Connect.Domains() as Promise<string[]>,
  },
  dashboard: {
    summary: () => Dashboard.Summary() as Promise<services.DashboardSummary>,
  },
  users: {
    list: (search: string, max: number) => list(Users.List(search, max)),
    get: (u: string) => one(Users.Get(u)),
    memberOf: (u: string) => list(Users.MemberOf(u)),
    licenseDetails: (u: string) => list(Users.LicenseDetails(u)),
    snapshot: (u: string) => one(Users.Snapshot(u)),
    block: (u: string) => Users.Block(u),
    unblock: (u: string) => Users.Unblock(u),
    resetPassword: (u: string, pw: string, force: boolean, confirm: string) =>
      Users.ResetPassword(u, pw, force, confirm),
    revokeSessions: (u: string, confirm: string) => Users.RevokeSessions(u, confirm),
    create: (name: string, upn: string, nick: string, pw: string, force: boolean, loc: string) =>
      one(Users.CreateUser(name, upn, nick, pw, force, loc)),
    inviteGuest: (email: string, name: string, redirectUrl: string, message: string, sendMail: boolean) =>
      Users.InviteGuest(email, name, redirectUrl, message, sendMail) as Promise<GraphObject>,
    update: (u: string, patchJSON: string) => Users.Update(u, patchJSON),
    setUsageLocation: (u: string, loc: string) => Users.SetUsageLocation(u, loc),
    delete: (u: string, confirm: string) => Users.Delete(u, confirm),
    getManager: (u: string) => one(Users.GetManager(u)),
    setManager: (u: string, mgr: string) => Users.SetManager(u, mgr),
    listDeleted: (max: number) => list(Users.ListDeleted(max)),
    restoreDeleted: (id: string) => one(Users.RestoreDeleted(id)),
  },
  authMethods: {
    list: (u: string) => list(AuthMethods.List(u)),
    resetMFA: (u: string, confirm: string) => AuthMethods.ResetMFA(u, confirm) as Promise<any>,
    createTAP: (u: string, lifetimeMinutes: number, oneTime: boolean) =>
      AuthMethods.CreateTAP(u, lifetimeMinutes, oneTime) as Promise<GraphObject>,
  },
  roles: {
    list: () => list(Roles.List()),
    members: (id: string) => list(Roles.Members(id)),
    addMember: (id: string, upn: string) => Roles.AddMember(id, upn),
    removeMember: (id: string, upn: string, confirm: string) => Roles.RemoveMember(id, upn, confirm),
  },
  licensing: {
    skus: () => list(Licensing.Skus()),
    assign: (u: string, add: string[], remove: string[]) => one(Licensing.Assign(u, add, remove)),
  },
  groups: {
    list: (search: string, max: number) => list(Groups.List(search, max)),
    get: (id: string) => one(Groups.Get(id)),
    members: (id: string) => list(Groups.Members(id)),
    addOwner: (id: string, upn: string) => Groups.AddOwner(id, upn),
    addMember: (id: string, upn: string) => Groups.AddMember(id, upn),
    createM365: (name: string, desc: string, nick: string, owner: string) =>
      one(Groups.CreateM365(name, desc, nick, owner)),
  },
  teams: {
    joined: (u: string) => list(Teams.JoinedTeams(u)),
    all: () => list(Teams.ListAllTeams(0)),
    channels: (t: string) => list(Teams.Channels(t)),
    teamMembers: (t: string) => list(Teams.TeamMembers(t)),
    channelMembers: (t: string, c: string) => list(Teams.ChannelMembers(t, c)),
    addTeamMember: (t: string, upn: string, owner: boolean) => one(Teams.AddTeamMember(t, upn, owner)),
    addChannelMember: (t: string, c: string, upn: string, owner: boolean) =>
      one(Teams.AddChannelMember(t, c, upn, owner)),
    removeTeamMember: (t: string, upn: string) => Teams.RemoveTeamMember(t, upn),
    removeChannelMember: (t: string, c: string, upn: string) => Teams.RemoveChannelMember(t, c, upn),
    teamify: (g: string) => one(Teams.Teamify(g)),
    createChannel: (t: string, name: string, desc: string, type: string, owner: string) =>
      one(Teams.CreateChannel(t, name, desc, type, owner)),
  },
  chats: {
    list: (u: string, max: number) => list(Chats.List(u, max)),
    messages: (c: string, top: number) => list(Chats.Messages(c, top)),
    members: (c: string) => list(Chats.Members(c)),
    listForPicker: (u: string) => Chats.ListForPicker(u) as Promise<{ id: string; label: string; chatType: string }[]>,
    addMember: (c: string, upn: string, owner: boolean) => one(Chats.AddMember(c, upn, owner)),
    removeMember: (c: string, upn: string) => Chats.RemoveMember(c, upn),
    createGroup: (topic: string, members: string[]) => one(Chats.CreateGroupChat(topic, members)),
  },
  mail: {
    list: (u: string, folder: string, top: number) => list(Mail.List(u, folder, top)),
    send: (u: string, subject: string, body: string, to: string[], confirm: string) =>
      Mail.Send(u, subject, body, to, confirm),
  },
  calendar: {
    list: (u: string, top: number) => list(Calendar.List(u, top)),
    createEvent: (
      u: string, subject: string, body: string, start: string, end: string, tz: string, attendees: string[],
    ) => one(Calendar.CreateEvent(u, subject, body, start, end, tz, attendees)),
  },
  drive: {
    sites: (search: string) => list(Drive.Sites(search)),
    sitesWithUsage: (search: string) => Drive.SitesWithUsage(search) as Promise<services.SiteUsage[]>,
    listRoot: (t: string, id: string) => list(Drive.ListRoot(t, id)),
    children: (t: string, id: string, item: string) => list(Drive.Children(t, id, item)),
    search: (t: string, id: string, q: string) => list(Drive.Search(t, id, q)),
    download: (t: string, id: string, item: string, name: string) => Drive.Download(t, id, item, name) as Promise<string>,
    upload: (t: string, id: string, folder: string) => one(Drive.Upload(t, id, folder)),
    delete: (t: string, id: string, item: string, confirm: string) => Drive.Delete(t, id, item, confirm),
    createLink: (t: string, id: string, item: string, type: string, scope: string) =>
      one(Drive.CreateLink(t, id, item, type, scope)),
    copyBetweenUsers: (src: string, dst: string, destFolder: string, overwrite: boolean) =>
      Drive.CopyBetweenUsers(src, dst, destFolder, overwrite) as Promise<services.CopyResult>,
    cancelTransfer: () => Drive.CancelTransfer(),
    offboardingPreview: (src: string) => Drive.OffboardingPreview(src) as Promise<services.CopyPreview>,
    quota: (user: string) => Drive.Quota(user) as Promise<services.DriveQuota>,
    pickTarget: (targets: string[], needed: number) => Drive.PickTarget(targets, needed) as Promise<string>,
    resumeCopy: (runId: string) => Drive.ResumeCopy(runId) as Promise<services.CopyResult>,
  },
  playbooks: {
    onboard: (req: any) => Playbook.Onboard(req) as Promise<services.PlaybookResult>,
    offboard: (req: any) => Playbook.Offboard(req) as Promise<services.PlaybookResult>,
    cancel: () => Playbook.Cancel() as Promise<void>,
  },
  access: {
    probe: () => Access.Probe() as Promise<Record<string, boolean>>,
  },
  auditQuery: {
    signIns: (upn: string, days: number, failedOnly: boolean, top: number) =>
      list(Audit.SignInsFiltered({ upn, days, failedOnly, top } as any)),
    directory: (search: string, days: number, top: number) => list(Audit.DirectoryAuditsFiltered(search, days, top)),
  },
  // "Give user1 the same access user2 has": diff first, then copy what is missing.
  mirror: {
    compare: (source: string, target: string) => Mirror.Compare(source, target) as Promise<services.AccessRow[]>,
    copy: (source: string, target: string, kinds: string[], confirm: string) =>
      Mirror.Copy({ source, target, kinds, confirm } as any) as Promise<services.PlaybookResult>,
    cancel: () => Mirror.Cancel() as Promise<void>,
  },
  intune: {
    devices: (max: number) => list(Intune.Devices(max)),
    device: (id: string) => one(Intune.Device(id)),
    wipe: (id: string, keepEnroll: boolean, keepUser: boolean, confirm: string) =>
      Intune.Wipe(id, keepEnroll, keepUser, confirm),
    retire: (id: string, confirm: string) => Intune.Retire(id, confirm),
    remoteLock: (id: string, confirm: string) => Intune.RemoteLock(id, confirm),
  },
  audit: {
    directory: (top: number) => list(Audit.DirectoryAudits(top)),
    signIns: (top: number) => list(Audit.SignIns(top)),
    activity: (n: number) => Audit.Activity(n) as Promise<any[]>,
  },
  raw: {
    send: (method: string, path: string, body: string) => one(Raw.Send(method, path, body)),
  },
  devices: {
    list: (search: string, max: number) => list(Devices.List(search, max)),
    get: (id: string) => one(Devices.Get(id)),
    enable: (id: string) => Devices.Enable(id),
    disable: (id: string) => Devices.Disable(id),
    delete: (id: string, confirm: string) => Devices.Delete(id, confirm),
    bitlockerKeys: (max: number) => list(Devices.BitLockerKeys(max)),
    bitlockerKeysForDevice: (deviceId: string) => list(Devices.BitLockerKeysForDevice(deviceId)),
    bitlockerKey: (id: string) => one(Devices.BitLockerKey(id)),
  },
  apps: {
    list: (search: string, max: number) => list(Apps.List(search, max)),
    expiring: (days: number) => Apps.Expiring(days) as Promise<services.ExpiringCredential[]>,
    addSecret: (objectId: string, name: string, months: number) =>
      Apps.AddSecret(objectId, name, months) as Promise<GraphObject>,
  },
  security: {
    caPolicies: () => list(Security.CAPolicies()),
    servicePrincipals: (search: string, max: number) => list(Security.ServicePrincipals(search, max)),
    oauthGrants: (spId: string) => list(Security.OAuthGrants(spId)),
    appRoleAssignments: (spId: string) => list(Security.AppRoleAssignments(spId)),
  },
  reports: {
    names: () => Reports.Names() as Promise<string[]>,
    csv: (report: string, period: string) => Reports.CSV(report, period) as Promise<string>,
  },
  cleanup: {
    findDuplicates: (ownerType: string, ownerId: string) => Cleanup.FindDuplicates(ownerType, ownerId) as Promise<services.DupGroup[]>,
    deleteItems: (refs: string[], confirm: string) => Cleanup.DeleteItems(refs, confirm) as Promise<any>,
    findVersionBloat: (ownerType: string, ownerId: string, minVersions: number, maxFiles: number) =>
      Cleanup.FindVersionBloat(ownerType, ownerId, minVersions, maxFiles) as Promise<services.VersionBloat[]>,
    trimVersions: (itemRef: string, keep: number, confirm: string) => Cleanup.TrimVersions(itemRef, keep, confirm) as Promise<any>,
    trimVersionsMany: (refs: string[], keep: number, confirm: string) =>
      Cleanup.TrimVersionsMany(refs, keep, confirm) as Promise<services.TrimResult[]>,
    cancelScan: () => Cleanup.CancelScan(),
  },
  health: {
    overview: () => list(ServiceHealth.Overview()),
    issues: () => list(ServiceHealth.Issues()),
    messages: () => list(ServiceHealth.Messages()),
  },
  update: {
    check: () => Update.Check() as Promise<services.UpdateInfo>,
    openReleases: (url: string) => Update.OpenReleasesPage(url),
    download: (assetUrl: string, name: string, size: number) => Update.Download(assetUrl, name, size) as Promise<string>,
    apply: (installerPath: string) => Update.Apply(installerPath) as Promise<void>,
  },
  journal: {
    list: (limit = 100) => Journal.List(limit) as Promise<journal.RunSummary[]>,
    get: (opId: string) => Journal.Get(opId) as Promise<journal.Run>,
  },
  notify: {
    get: () => Notify.Get() as Promise<services.NotifyConfig>,
    set: (cfg: services.NotifyConfig) => Notify.Set(cfg) as Promise<void>,
    test: () => Notify.Test() as Promise<void>,
  },
}

export function errParsed(e: unknown): ParsedError {
  let msg = ''
  if (e == null) msg = ''
  else if (typeof e === 'string') msg = e
  else if (e instanceof Error) msg = e.message
  else msg = String(e)
  return parseErr(msg)
}

export function errMessage(e: unknown): string {
  const p = errParsed(e)
  let msg = p.message
  if (p.code && p.status) msg = `${p.status} ${p.code}: ${msg}`
  if (p.hint) {
    msg += ' — ' + i18n.t('errors.grantPermission', { p: p.hint })
  } else if (p.status === 403) {
    // Missing-permission 403 without a known mapping: still point the way.
    msg += ' — ' + i18n.t('errors.missingPermission')
  }
  if (p.requestId) msg += ` (requestId=${p.requestId})`
  return msg
}
