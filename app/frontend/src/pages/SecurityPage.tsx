import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ShieldAlert, AppWindow, Search } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Input, Badge, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

type Tab = 'ca' | 'consents'

const stateBadge: Record<string, 'ok' | 'warn' | 'neutral'> = {
  enabled: 'ok',
  enabledForReportingButNotEnforced: 'warn',
  disabled: 'neutral',
}

export function SecurityPage() {
  const { t } = useTranslation()
  const { connected, toast, cache, setCache } = useStore()
  const [tab, setTab] = useState<Tab>('ca')

  const [policies, setPolicies] = useState<GraphObject[] | null>(null)
  const [selPolicy, setSelPolicy] = useState<GraphObject | null>(null)

  // Cache-backed: keep the SP search results and the selected SP's consents
  // across navigation so leaving the page does not force a re-search.
  const [search, setSearchLocal] = useState<string>(() => cache['security.spSearch'] ?? '')
  const setSearch = (v: string) => { setSearchLocal(v); setCache('security.spSearch', v) }
  const [sps, setSpsLocal] = useState<GraphObject[] | null>(() => cache['security.sps'] ?? null)
  const setSps = (v: GraphObject[] | null) => { setSpsLocal(v); setCache('security.sps', v) }
  const [selSp, setSelSpLocal] = useState<GraphObject | null>(() => cache['security.selSp'] ?? null)
  const setSelSp = (v: GraphObject | null) => { setSelSpLocal(v); setCache('security.selSp', v) }
  const [grants, setGrantsLocal] = useState<GraphObject[] | null>(() => cache['security.grants'] ?? null)
  const setGrants = (v: GraphObject[] | null) => { setGrantsLocal(v); setCache('security.grants', v) }
  const [appRoles, setAppRolesLocal] = useState<GraphObject[] | null>(() => cache['security.appRoles'] ?? null)
  const setAppRoles = (v: GraphObject[] | null) => { setAppRolesLocal(v); setCache('security.appRoles', v) }
  const [loading, setLoading] = useState(false)

  useEffect(() => {
    if (!connected || tab !== 'ca' || policies) return
    api.security.caPolicies().then(setPolicies).catch((e) => toast('err', errMessage(e)))
  }, [connected, tab])

  const loadSps = () => {
    setLoading(true)
    api.security.servicePrincipals(search, 500)
      .then(setSps).catch((e) => toast('err', errMessage(e)))
      .finally(() => setLoading(false))
  }

  const pickSp = (sp: GraphObject) => {
    setSelSp(sp); setGrants(null); setAppRoles(null)
    api.security.oauthGrants(sp.id).then(setGrants).catch(() => setGrants([]))
    api.security.appRoleAssignments(sp.id).then(setAppRoles).catch(() => setAppRoles([]))
  }

  return (
    <Page title={t('security.title')} subtitle={t('security.subtitle')}>
      <div className="mb-4 inline-flex rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
        {([['ca', <ShieldAlert size={15} key="c" />, t('security.tabCa')], ['consents', <AppWindow size={15} key="a" />, t('security.tabConsents')]] as const).map(([m, icon, label]) => (
          <button key={m} onClick={() => setTab(m)}
            className={`flex items-center gap-2 rounded-md px-4 py-1.5 text-sm font-medium ${tab === m ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)]'}`}>
            {icon}{label}
          </button>
        ))}
      </div>

      {tab === 'ca' && (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-[minmax(320px,420px)_1fr]">
          <Card title={t('security.tabCa')}>
            {!policies && <Spinner />}
            {policies && policies.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
            <div className="flex flex-col gap-1.5">
              {(policies || []).map((p) => (
                <button key={p.id} onClick={() => setSelPolicy(p)}
                  className={`flex items-center gap-2 rounded-lg border px-3 py-2 text-left text-sm ${selPolicy?.id === p.id ? 'border-[var(--accent)] bg-[var(--bg-elev)]' : 'border-[var(--border)] bg-[var(--bg)]'}`}>
                  <span className="min-w-0 flex-1 truncate">{p.displayName}</span>
                  <Badge kind={stateBadge[p.state] || 'neutral'}>{t(`security.state.${p.state}`, { defaultValue: String(p.state) })}</Badge>
                </button>
              ))}
            </div>
          </Card>
          <Card title={selPolicy?.displayName || t('security.detail')}>
            {!selPolicy && <p className="text-sm text-[var(--text-faint)]">{t('security.pickPolicy')}</p>}
            {selPolicy && (
              <pre className="max-h-[70vh] overflow-auto rounded-lg bg-[var(--bg)] p-3 font-mono text-xs leading-relaxed text-[var(--text-dim)]">
                {JSON.stringify(selPolicy, null, 2)}
              </pre>
            )}
          </Card>
        </div>
      )}

      {tab === 'consents' && (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-[minmax(320px,420px)_1fr]">
          <Card title={t('security.tabConsents')}>
            <div className="mb-2 flex gap-2">
              <Input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={t('common.search')}
                onKeyDown={(e) => e.key === 'Enter' && loadSps()} />
              <Button variant="primary" onClick={loadSps} disabled={loading}>{loading ? <Spinner /> : <Search size={15} />}</Button>
            </div>
            {sps && sps.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
            <div className="flex max-h-[65vh] flex-col gap-1.5 overflow-y-auto">
              {(sps || []).map((sp) => (
                <button key={sp.id} onClick={() => pickSp(sp)}
                  className={`flex items-center gap-2 rounded-lg border px-3 py-2 text-left text-sm ${selSp?.id === sp.id ? 'border-[var(--accent)] bg-[var(--bg-elev)]' : 'border-[var(--border)] bg-[var(--bg)]'}`}>
                  <span className="min-w-0 flex-1 truncate">{sp.displayName}</span>
                  {sp.accountEnabled === false && <Badge kind="warn">{t('security.disabledSp')}</Badge>}
                </button>
              ))}
            </div>
          </Card>
          <Card title={selSp?.displayName || t('security.detail')}>
            {!selSp && <p className="text-sm text-[var(--text-faint)]">{t('security.pickSp')}</p>}
            {selSp && (
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
          </Card>
        </div>
      )}
    </Page>
  )
}
