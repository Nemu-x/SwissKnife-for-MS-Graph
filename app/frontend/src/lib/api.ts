// Typed wrapper over the generated Wails bindings.
// json.RawMessage arrives from Go as real JSON (mistyped as number[] in .d.ts),
// so we cast it to proper types here.
import * as Connect from '../../wailsjs/go/services/ConnectService'
import * as Dashboard from '../../wailsjs/go/services/DashboardService'
import * as Playbook from '../../wailsjs/go/services/PlaybookService'
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
import * as Update from '../../wailsjs/go/services/UpdateService'
import type { secrets, services } from '../../wailsjs/go/models'

export type GraphObject = Record<string, any>
export type Profile = secrets.Profile
export type Status = services.Status
export type ConnectRequest = services.ConnectRequest

const list = (p: Promise<any>) => p as Promise<GraphObject[]>
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
    offboardingPreview: (src: string) => Drive.OffboardingPreview(src) as Promise<services.CopyPreview>,
    quota: (user: string) => Drive.Quota(user) as Promise<services.DriveQuota>,
    pickTarget: (targets: string[], needed: number) => Drive.PickTarget(targets, needed) as Promise<string>,
  },
  playbooks: {
    onboard: (req: any) => Playbook.Onboard(req) as Promise<services.PlaybookResult>,
    offboard: (req: any) => Playbook.Offboard(req) as Promise<services.PlaybookResult>,
  },
  access: {
    probe: () => Access.Probe() as Promise<Record<string, boolean>>,
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
    bitlockerKey: (id: string) => one(Devices.BitLockerKey(id)),
  },
  apps: {
    list: (search: string, max: number) => list(Apps.List(search, max)),
    expiring: (days: number) => Apps.Expiring(days) as Promise<services.ExpiringCredential[]>,
  },
  reports: {
    names: () => Reports.Names() as Promise<string[]>,
    csv: (report: string, period: string) => Reports.CSV(report, period) as Promise<string>,
  },
  cleanup: {
    findDuplicates: (ownerType: string, ownerId: string) => Cleanup.FindDuplicates(ownerType, ownerId) as Promise<services.DupGroup[]>,
    deleteItems: (ownerType: string, ownerId: string, ids: string[], confirm: string) =>
      Cleanup.DeleteItems(ownerType, ownerId, ids, confirm) as Promise<any>,
  },
  health: {
    overview: () => list(ServiceHealth.Overview()),
    issues: () => list(ServiceHealth.Issues()),
    messages: () => list(ServiceHealth.Messages()),
  },
  update: {
    check: () => Update.Check() as Promise<services.UpdateInfo>,
    openReleases: (url: string) => Update.OpenReleasesPage(url),
  },
}

export function errMessage(e: unknown): string {
  let msg = ''
  if (e == null) msg = ''
  else if (typeof e === 'string') msg = e
  else if (e instanceof Error) msg = e.message
  else msg = String(e)
  // Enrich the common "missing permission" case so it isn't a cryptic 403.
  if (/\b403\b/.test(msg)) {
    msg += ' — the app registration is likely missing a Graph permission for this call. Add the required Application permission and grant admin consent.'
  }
  return msg
}
