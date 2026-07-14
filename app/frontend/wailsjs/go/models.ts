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
	export class CopyResult {
	    copied: string[];
	    skipped: Record<string, string>;
	    failed: Record<string, string>;
	
	    static createFrom(source: any = {}) {
	        return new CopyResult(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.copied = source["copied"];
	        this.skipped = source["skipped"];
	        this.failed = source["failed"];
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

}

