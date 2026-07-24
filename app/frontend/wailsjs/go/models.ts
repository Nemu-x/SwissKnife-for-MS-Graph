export namespace auditlog {
	
	export class Entry {
	    // Go type: time
	    time: any;
	    action: string;
	    target: string;
	    detail?: string;
	    ok: boolean;
	    error?: string;
	    profile?: string;
	
	    static createFrom(source: any = {}) {
	        return new Entry(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.time = this.convertValues(source["time"], null);
	        this.action = source["action"];
	        this.target = source["target"];
	        this.detail = source["detail"];
	        this.ok = source["ok"];
	        this.error = source["error"];
	        this.profile = source["profile"];
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}

}

export namespace secrets {
	
	export class Profile {
	    id: string;
	    name: string;
	    tenantId: string;
	    clientId: string;
	    authMode: string;
	    hasSecret: boolean;
	
	    static createFrom(source: any = {}) {
	        return new Profile(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.id = source["id"];
	        this.name = source["name"];
	        this.tenantId = source["tenantId"];
	        this.clientId = source["clientId"];
	        this.authMode = source["authMode"];
	        this.hasSecret = source["hasSecret"];
	    }
	}

}

export namespace services {
	
	export class ChannelRef {
	    teamId: string;
	    channelId: string;
	
	    static createFrom(source: any = {}) {
	        return new ChannelRef(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.teamId = source["teamId"];
	        this.channelId = source["channelId"];
	    }
	}
	export class ChatPickItem {
	    id: string;
	    label: string;
	    chatType: string;
	
	    static createFrom(source: any = {}) {
	        return new ChatPickItem(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.id = source["id"];
	        this.label = source["label"];
	        this.chatType = source["chatType"];
	    }
	}
	export class ConnectRequest {
	    profileId: string;
	    tenantId: string;
	    clientId: string;
	    secret: string;
	    authMode: string;
	    rememberAs: string;
	
	    static createFrom(source: any = {}) {
	        return new ConnectRequest(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.profileId = source["profileId"];
	        this.tenantId = source["tenantId"];
	        this.clientId = source["clientId"];
	        this.secret = source["secret"];
	        this.authMode = source["authMode"];
	        this.rememberAs = source["rememberAs"];
	    }
	}
	export class CopyPreview {
	    files: number;
	    folders: number;
	    totalBytes: number;
	
	    static createFrom(source: any = {}) {
	        return new CopyPreview(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.files = source["files"];
	        this.folders = source["folders"];
	        this.totalBytes = source["totalBytes"];
	    }
	}
	export class CopyResult {
	    copied: string[];
	    skipped: Record<string, string>;
	    failed: Record<string, string>;
	    canceled: boolean;
	
	    static createFrom(source: any = {}) {
	        return new CopyResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.copied = source["copied"];
	        this.skipped = source["skipped"];
	        this.failed = source["failed"];
	        this.canceled = source["canceled"];
	    }
	}
	export class LicenseLine {
	    skuPartNumber: string;
	    consumed: number;
	    total: number;
	
	    static createFrom(source: any = {}) {
	        return new LicenseLine(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.skuPartNumber = source["skuPartNumber"];
	        this.consumed = source["consumed"];
	        this.total = source["total"];
	    }
	}
	export class DashboardSummary {
	    orgName: string;
	    users: number;
	    groups: number;
	    domains: number;
	    licensesUsed: number;
	    licensesTotal: number;
	    licenses: LicenseLine[];
	
	    static createFrom(source: any = {}) {
	        return new DashboardSummary(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.orgName = source["orgName"];
	        this.users = source["users"];
	        this.groups = source["groups"];
	        this.domains = source["domains"];
	        this.licensesUsed = source["licensesUsed"];
	        this.licensesTotal = source["licensesTotal"];
	        this.licenses = this.convertValues(source["licenses"], LicenseLine);
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}
	export class DriveQuota {
	    total: number;
	    used: number;
	    remaining: number;
	
	    static createFrom(source: any = {}) {
	        return new DriveQuota(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.total = source["total"];
	        this.used = source["used"];
	        this.remaining = source["remaining"];
	    }
	}
	export class DupItem {
	    ref: string;
	    name: string;
	    path: string;
	
	    static createFrom(source: any = {}) {
	        return new DupItem(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.ref = source["ref"];
	        this.name = source["name"];
	        this.path = source["path"];
	    }
	}
	export class DupGroup {
	    name: string;
	    size: number;
	    count: number;
	    wasted: number;
	    items: DupItem[];
	
	    static createFrom(source: any = {}) {
	        return new DupGroup(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.name = source["name"];
	        this.size = source["size"];
	        this.count = source["count"];
	        this.wasted = source["wasted"];
	        this.items = this.convertValues(source["items"], DupItem);
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}
	
	export class ExpiringCredential {
	    appName: string;
	    appId: string;
	    kind: string;
	    displayName: string;
	    expires: string;
	    daysLeft: number;
	
	    static createFrom(source: any = {}) {
	        return new ExpiringCredential(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.appName = source["appName"];
	        this.appId = source["appId"];
	        this.kind = source["kind"];
	        this.displayName = source["displayName"];
	        this.expires = source["expires"];
	        this.daysLeft = source["daysLeft"];
	    }
	}
	
	export class OffboardRequest {
	    upn: string;
	    confirm: string;
	    block: boolean;
	    revokeSessions: boolean;
	    oof: boolean;
	    oofMessage: string;
	    forwardTo: string;
	    hideFromGal: boolean;
	    calendarTo: string;
	    removeFromGroups: boolean;
	    removeAllLicenses: boolean;
	    backupToUser: string;
	    backupFolder: string;
	    delete: boolean;
	
	    static createFrom(source: any = {}) {
	        return new OffboardRequest(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.upn = source["upn"];
	        this.confirm = source["confirm"];
	        this.block = source["block"];
	        this.revokeSessions = source["revokeSessions"];
	        this.oof = source["oof"];
	        this.oofMessage = source["oofMessage"];
	        this.forwardTo = source["forwardTo"];
	        this.hideFromGal = source["hideFromGal"];
	        this.calendarTo = source["calendarTo"];
	        this.removeFromGroups = source["removeFromGroups"];
	        this.removeAllLicenses = source["removeAllLicenses"];
	        this.backupToUser = source["backupToUser"];
	        this.backupFolder = source["backupFolder"];
	        this.delete = source["delete"];
	    }
	}
	export class OnboardRequest {
	    displayName: string;
	    upn: string;
	    mailNickname: string;
	    password: string;
	    usageLocation: string;
	    skuIds: string[];
	    groupIds: string[];
	    teamIds: string[];
	    channelRefs: ChannelRef[];
	
	    static createFrom(source: any = {}) {
	        return new OnboardRequest(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.displayName = source["displayName"];
	        this.upn = source["upn"];
	        this.mailNickname = source["mailNickname"];
	        this.password = source["password"];
	        this.usageLocation = source["usageLocation"];
	        this.skuIds = source["skuIds"];
	        this.groupIds = source["groupIds"];
	        this.teamIds = source["teamIds"];
	        this.channelRefs = this.convertValues(source["channelRefs"], ChannelRef);
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}
	export class Step {
	    name: string;
	    nameKey?: string;
	    ok: boolean;
	    detail?: string;
	    detailKey?: string;
	    params?: Record<string, any>;
	    error?: string;
	    errorCode?: string;
	    hint?: string;
	
	    static createFrom(source: any = {}) {
	        return new Step(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.name = source["name"];
	        this.nameKey = source["nameKey"];
	        this.ok = source["ok"];
	        this.detail = source["detail"];
	        this.detailKey = source["detailKey"];
	        this.params = source["params"];
	        this.error = source["error"];
	        this.errorCode = source["errorCode"];
	        this.hint = source["hint"];
	    }
	}
	export class PlaybookResult {
	    ok: boolean;
	    canceled: boolean;
	    steps: Step[];
	
	    static createFrom(source: any = {}) {
	        return new PlaybookResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.ok = source["ok"];
	        this.canceled = source["canceled"];
	        this.steps = this.convertValues(source["steps"], Step);
	    }
	
		convertValues(a: any, classs: any, asMap: boolean = false): any {
		    if (!a) {
		        return a;
		    }
		    if (a.slice && a.map) {
		        return (a as any[]).map(elem => this.convertValues(elem, classs));
		    } else if ("object" === typeof a) {
		        if (asMap) {
		            for (const key of Object.keys(a)) {
		                a[key] = new classs(a[key]);
		            }
		            return a;
		        }
		        return new classs(a);
		    }
		    return a;
		}
	}
	export class SiteUsage {
	    id: string;
	    name: string;
	    webUrl: string;
	    used: number;
	
	    static createFrom(source: any = {}) {
	        return new SiteUsage(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.id = source["id"];
	        this.name = source["name"];
	        this.webUrl = source["webUrl"];
	        this.used = source["used"];
	    }
	}
	export class Status {
	    connected: boolean;
	    profileName: string;
	    readOnly: boolean;
	    org?: number[];
	
	    static createFrom(source: any = {}) {
	        return new Status(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.connected = source["connected"];
	        this.profileName = source["profileName"];
	        this.readOnly = source["readOnly"];
	        this.org = source["org"];
	    }
	}
	
	export class TrimResult {
	    ref: string;
	    removed: number;
	    error?: string;
	
	    static createFrom(source: any = {}) {
	        return new TrimResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.ref = source["ref"];
	        this.removed = source["removed"];
	        this.error = source["error"];
	    }
	}
	export class UpdateInfo {
	    currentVersion: string;
	    latestVersion: string;
	    updateAvailable: boolean;
	    notes: string;
	    url: string;
	
	    static createFrom(source: any = {}) {
	        return new UpdateInfo(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.currentVersion = source["currentVersion"];
	        this.latestVersion = source["latestVersion"];
	        this.updateAvailable = source["updateAvailable"];
	        this.notes = source["notes"];
	        this.url = source["url"];
	    }
	}
	export class VersionBloat {
	    ref: string;
	    name: string;
	    path: string;
	    versions: number;
	    currentSize: number;
	    reclaimable: number;
	
	    static createFrom(source: any = {}) {
	        return new VersionBloat(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.ref = source["ref"];
	        this.name = source["name"];
	        this.path = source["path"];
	        this.versions = source["versions"];
	        this.currentSize = source["currentSize"];
	        this.reclaimable = source["reclaimable"];
	    }
	}

}

