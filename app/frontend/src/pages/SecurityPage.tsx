import { useCallback, useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ShieldAlert, AppWindow, Search } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { Button, Field, Input, Badge, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

type Tab = 'ca' | 'consents'

const stateBadge: Record<string, 'ok' | 'warn' | 'neutral'> = {
  enabled: 'ok',
  enabledForReportingButNotEnforced: 'warn',
  disabled: 'neutral',
}

// Conditional Access policies are deep JSON; these pull out the handful of
// facts an operator checks: who it hits, what it covers, what it demands.
function count(v: unknown): number { return Array.isArray(v) ? v.length : 0 }

function targetSummary(t: (k: string, o?: any) => string, users: any): string[] {
  if (!users) return []
  const out: string[] = []
  const inc = users.includeUsers || []
  if (inc.includes('All')) out.push(t('security.allUsers'))
  else if (inc.includes('GuestsOrExternalUsers')) out.push(t('security.guests'))
  else if (count(inc)) out.push(t('security.nUsers', { n: count(inc) }))
  if (count(users.includeGroups)) out.push(t('security.nGroups', { n: count(users.includeGroups) }))
  if (count(users.includeRoles)) out.push(t('security.nRoles', { n: count(users.includeRoles) }))
  return out
}

function excludeSummary(t: (k: string, o?: any) => string, users: any): string[] {
  if (!users) return []
  const out: string[] = []
  if (count(users.excludeUsers)) out.push(t('security.nUsers', { n: count(users.excludeUsers) }))
  if (count(users.excludeGroups)) out.push(t('security.nGroups', { n: count(users.excludeGroups) }))
  if (count(users.excludeRoles)) out.push(t('security.nRoles', { n: count(users.excludeRoles) }))
  return out
}

export function SecurityPage() {
  const { t } = useTranslation()
  const { toast, cache, setCache } = useStore()
  const [tab, setTab] = useState<Tab>('ca')

  const [policies, setPolicies] = useState<GraphObject[] | null>(() => cache['security.policies'] ?? null)
  const [selPolicy, setSelPolicy] = useState<GraphObject | null>(null)
  const [loadingCa, setLoadingCa] = useState(false)

  const [search, setSearchLocal] = useState<string>(() => cache['security.spSearch'] ?? '')
  const setSearch = (v: string) => { setSearchLocal(v); setCache('security.spSearch', v) }
  const [sps, setSpsLocal] = useState<GraphObject[] | null>(() => cache['security.sps'] ?? null)
  const setSps = (v: GraphObject[] | null) => { setSpsLocal(v); setCache('security.sps', v) }
  const [selSp, setSelSpLocal] = useState<GraphObject | null>(() => cache['security.selSp'] ?? null)
  const setSelSp = (v: GraphObject | null) => { setSelSpLocal(v); setCache('security.selSp', v) }
  const [grants, setGrants] = useState<GraphObject[] | null>(null)
  const [appRoles, setAppRoles] = useState<GraphObject[] | null>(null)
  const [loadingSp, setLoadingSp] = useState(false)
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  const loadPolicies = () => {
    setTab('ca'); setLoadingCa(true); setSelPolicy(null)
    api.security.caPolicies()
      .then((p) => {
        setPolicies(p); setCache('security.policies', p)
        const on = (p || []).filter((x) => x.state === 'enabled').length
        setStatus((s) => ({ ...s, ca: { ok: true, text: t('security.caSummary', { n: (p || []).length, on }), at: Date.now() } }))
      })
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setLoadingCa(false))
  }

  const loadSps = () => {
    setTab('consents'); setLoadingSp(true); setSelSp(null)
    api.security.servicePrincipals(search, 500)
      .then((r) => {
        setSps(r)
        setStatus((s) => ({ ...s, consents: { ok: true, text: t('security.spsFound', { n: (r || []).length }), at: Date.now() } }))
      })
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setLoadingSp(false))
  }

  const loadSpDetail = useCallback((sp: GraphObject) => {
    api.security.oauthGrants(sp.id).then(setGrants).catch(() => setGrants([]))
    api.security.appRoleAssignments(sp.id).then(setAppRoles).catch(() => setAppRoles([]))
  }, [])

  const pickSp = (sp: GraphObject) => {
    setSelSp(sp); setGrants(null); setAppRoles(null)
    loadSpDetail(sp)
  }

  // The selection is restored from the cache on mount, but its grants are not —
  // without this the detail pane spins forever after navigating back.
  useEffect(() => {
    if (selSp && !grants && !appRoles) loadSpDetail(selSp)
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount-time rehydration only
  }, [])

  const actions: TaskAction[] = [
    {
      id: 'ca', label: t('security.tileCa'), hint: t('security.hintCa'), icon: <ShieldAlert size={16} />, variant: 'primary',
      note: <p>{t('security.noteCa')}</p>,
      panel: (
        <TaskForm>
          <Button variant="primary" disabled={loadingCa} onClick={loadPolicies}>
            {loadingCa ? <Spinner /> : <ShieldAlert size={15} />} {t('security.loadPolicies')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'consents', label: t('security.tileConsents'), hint: t('security.hintConsents'), icon: <AppWindow size={16} />, variant: 'primary',
      note: <p>{t('security.noteConsents')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('common.search')} hint={t('security.searchHint')}>
            <Input value={search} onChange={(e) => setSearch(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && loadSps()} />
          </Field>
          <Button variant="primary" disabled={loadingSp} onClick={loadSps}>
            {loadingSp ? <Spinner /> : <Search size={15} />} {t('common.search')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  const policyDetail = (p: GraphObject) => {
    const c = p.conditions || {}
    const g = p.grantControls || {}
    const controls: string[] = (g.builtInControls || []).map((b: string) => t(`security.control.${b}`, { defaultValue: b }))
    const apps = c.applications || {}
    const includeApps: string[] = apps.includeApplications || []
    const rows: [string, string][] = [
      [t('security.appliesTo'), targetSummary(t, c.users).join(' · ') || '—'],
      [t('security.excludes'), excludeSummary(t, c.users).join(' · ') || '—'],
      [t('security.apps'), includeApps.includes('All') ? t('security.allApps') : includeApps.length ? t('security.nApps', { n: includeApps.length }) : '—'],
      [t('security.clientApps'), (c.clientAppTypes || []).join(', ') || '—'],
      [t('security.platforms'), count(c.platforms?.includePlatforms) ? (c.platforms.includePlatforms as string[]).join(', ') : '—'],
      [t('security.locations'), count(c.locations?.includeLocations) ? t('security.nLocations', { n: count(c.locations.includeLocations) }) : '—'],
      [t('security.requires'), controls.length ? `${controls.join(g.operator === 'OR' ? ' / ' : ' + ')}` : '—'],
      [t('security.risk'), [...(c.signInRiskLevels || []), ...(c.userRiskLevels || [])].join(', ') || '—'],
    ]
    const blocking = (g.builtInControls || []).includes('block')
    return (
      <div className="flex flex-col gap-3">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold">{p.displayName}</span>
          <Badge kind={stateBadge[p.state] || 'neutral'}>{t(`security.state.${p.state}`, { defaultValue: String(p.state) })}</Badge>
          {blocking && <Badge kind="danger">{t('security.blocks')}</Badge>}
        </div>
        <dl className="flex flex-col">
          {rows.map(([k, v]) => (
            <div key={k} className="grid grid-cols-[150px_1fr] gap-3 border-b border-[var(--border)]/50 py-1.5 last:border-0">
              <dt className="text-xs uppercase tracking-wide text-[var(--text-faint)]">{k}</dt>
              <dd className="text-sm text-[var(--text)]">{v}</dd>
            </div>
          ))}
        </dl>
        <details className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-2">
          <summary className="cursor-pointer text-xs text-[var(--text-dim)]">{t('security.rawJson')}</summary>
          <pre className="mt-2 max-h-72 overflow-auto text-xs leading-relaxed text-[var(--text-dim)]">{JSON.stringify(p, null, 2)}</pre>
        </details>
      </div>
    )
  }

  const resultPane = (
    <div className="grid h-full grid-cols-1 lg:grid-cols-[minmax(240px,320px)_1fr]">
      <div className="min-h-0 overflow-auto border-b border-[var(--border)] p-2 lg:border-b-0 lg:border-r">
        {tab === 'ca' && (policies || []).map((p) => (
          <button key={p.id} onClick={() => setSelPolicy(p)}
            className={`mb-1 flex w-full items-center gap-2 rounded-lg border px-3 py-2 text-left text-sm ${selPolicy?.id === p.id ? 'border-[var(--accent)] bg-[var(--accent)]/10' : 'border-transparent hover:bg-[var(--bg-elev-2)]'}`}>
            <span className="min-w-0 flex-1 truncate">{p.displayName}</span>
            <Badge kind={stateBadge[p.state] || 'neutral'}>{t(`security.state.${p.state}`, { defaultValue: String(p.state) })}</Badge>
          </button>
        ))}
        {tab === 'consents' && (sps || []).map((sp) => (
          <button key={sp.id} onClick={() => pickSp(sp)}
            className={`mb-1 flex w-full items-center gap-2 rounded-lg border px-3 py-2 text-left text-sm ${selSp?.id === sp.id ? 'border-[var(--accent)] bg-[var(--accent)]/10' : 'border-transparent hover:bg-[var(--bg-elev-2)]'}`}>
            <span className="min-w-0 flex-1 truncate">{sp.displayName}</span>
            {sp.accountEnabled === false && <Badge kind="warn">{t('security.disabledSp')}</Badge>}
          </button>
        ))}
        {tab === 'ca' && policies && policies.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
        {tab === 'consents' && sps && sps.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
      </div>

      <div className="min-h-0 overflow-auto p-4">
        {tab === 'ca' && !selPolicy && <p className="text-sm text-[var(--text-faint)]">{t('security.pickPolicy')}</p>}
        {tab === 'ca' && selPolicy && policyDetail(selPolicy)}

        {tab === 'consents' && !selSp && <p className="text-sm text-[var(--text-faint)]">{t('security.pickSp')}</p>}
        {tab === 'consents' && selSp && (
          <div className="flex flex-col gap-4">
            <div className="text-xs text-[var(--text-faint)]">appId: <span className="font-mono">{selSp.appId}</span></div>
            <div>
              <div className="mb-1 text-sm font-medium">{t('security.delegated')}</div>
              {!grants && <Spinner />}
              {grants && grants.length === 0 && <p className="text-xs text-[var(--text-faint)]">{t('common.empty')}</p>}
              <div className="flex flex-col gap-1">
                {(grants || []).map((g, i) => (
                  <div key={i} className="rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs">
                    <span className="text-[var(--text-dim)]">{g.consentType === 'AllPrincipals' ? t('security.adminConsent') : t('security.userConsent')}:</span>{' '}
                    <span className="font-mono">{(g.scope || '').trim() || '—'}</span>
                  </div>
                ))}
              </div>
            </div>
            <div>
              <div className="mb-1 text-sm font-medium">{t('security.application')}</div>
              {!appRoles && <Spinner />}
              {appRoles && appRoles.length === 0 && <p className="text-xs text-[var(--text-faint)]">{t('common.empty')}</p>}
              <div className="flex flex-col gap-1">
                {(appRoles || []).map((a, i) => (
                  <div key={i} className="rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs">
                    <span className="text-[var(--text-dim)]">{a.resourceDisplayName}:</span>{' '}
                    <span className="font-mono">{a.appRoleId}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )

  return (
    <TaskPage
      pageId="security"
      title={t('security.title')}
      subtitle={t('security.subtitle')}
      actions={actions}
      status={status}
      busy={loadingCa || loadingSp}
      busyLabel={loadingCa ? t('security.loadPolicies') : t('common.search')}
      hasResult={(tab === 'ca' && !!policies) || (tab === 'consents' && !!sps) || loadingCa || loadingSp}
      onClearResult={() => { setPolicies(null); setSelPolicy(null); setSps(null); setSelSp(null); setCache('security.policies', null); setCache('security.sps', null) }}
      result={resultPane}
    />
  )
}
