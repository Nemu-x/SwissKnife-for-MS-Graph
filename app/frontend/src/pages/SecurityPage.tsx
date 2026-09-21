import { useCallback, useEffect, useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ShieldAlert, AppWindow, Search, Lightbulb, Camera, GitCompare, Trash2, ExternalLink } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select, Badge, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { useConfirm } from '../lib/useConfirm'
import { api, errMessage, type GraphObject } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

type Tab = 'ca' | 'consents' | 'recommendations' | 'snapshot' | 'drift'
type SnapshotMeta = services.SnapshotMeta
type SnapshotDiff = services.SnapshotDiff
type SectionDiff = services.SectionDiff
type ObjectChange = services.ObjectChange
type BadgeKind = 'ok' | 'warn' | 'danger' | 'neutral'

const stateBadge: Record<string, 'ok' | 'warn' | 'neutral'> = {
  enabled: 'ok',
  enabledForReportingButNotEnforced: 'warn',
  disabled: 'neutral',
}

// Recommendations: the ones still open and most urgent come first.
const PRIORITY_ORDER: Record<string, number> = { critical: 0, high: 1, medium: 2, low: 3 }
const priorityBadge: Record<string, BadgeKind> = { critical: 'danger', high: 'danger', medium: 'warn', low: 'neutral' }
function statusBadge(status: string): BadgeKind {
  if (status === 'active' || status === 'needsMoreAction') return 'warn'
  if (status === 'completedBySystem' || status === 'completedByUser') return 'ok'
  return 'neutral'
}
function sortRecs(list: GraphObject[]): GraphObject[] {
  const open = (r: GraphObject) => (r.status === 'active' || r.status === 'needsMoreAction' ? 0 : 1)
  return [...list].sort((a, b) =>
    open(a) - open(b) || (PRIORITY_ORDER[a.priority] ?? 9) - (PRIORITY_ORDER[b.priority] ?? 9) || String(a.displayName).localeCompare(String(b.displayName)))
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

// A diff value: scalars inline, structures as pretty JSON.
function DiffValue({ v }: { v: any }) {
  if (v === null || v === undefined) return <span className="text-[var(--text-faint)]">—</span>
  if (typeof v !== 'object') return <span className="break-words">{String(v)}</span>
  return <pre className="max-h-40 overflow-auto whitespace-pre-wrap break-all text-[11px] leading-snug">{JSON.stringify(v, null, 1)}</pre>
}

type ImpactedState = GraphObject[] | 'loading' | { error: string }

// "section|done|total" as the store records it from snapshot:progress events.
function parseSnapshotProgress(p?: string): { section: string; done: number; total: number } | null {
  if (!p) return null
  const [section, done, total] = p.split('|')
  return section ? { section, done: Number(done) || 0, total: Number(total) || 0 } : null
}

export function SecurityPage() {
  const { t, i18n } = useTranslation()
  const { toast, cache, setCache, jobs, patchJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
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

  // Entra recommendations: one expensive list, cached across navigation.
  const [recs, setRecsLocal] = useState<GraphObject[] | null>(() => cache['security.recommendations'] ?? null)
  const setRecs = (v: GraphObject[] | null) => { setRecsLocal(v); setCache('security.recommendations', v) }
  const [selRec, setSelRec] = useState<GraphObject | null>(null)
  const [impacted, setImpacted] = useState<Record<string, ImpactedState>>({})
  const [loadingRecs, setLoadingRecs] = useState(false)
  const [rawRecs, setRawRecs] = useState(false)

  // Snapshots live on disk, so the list is loaded on mount regardless of the
  // Graph connection; taking one needs the connection.
  const [snapName, setSnapName] = useState('')
  const [snaps, setSnaps] = useState<SnapshotMeta[] | null>(null)
  const [selSnap, setSelSnap] = useState<SnapshotMeta | null>(null)
  const [snapRaw, setSnapRaw] = useState<any | null>(null)
  // The crawl runs as the "snapshot" store job: leaving the page keeps it
  // going, and the result is waiting in the job when the operator comes back.
  const snapJob = jobs.snapshot
  const taking = !!snapJob?.running
  const progress = parseSnapshotProgress(snapJob?.progress)
  const [diffA, setDiffA] = useState('')
  const [diffB, setDiffB] = useState('')
  const [diff, setDiff] = useState<SnapshotDiff | null>(null)
  const [diffing, setDiffing] = useState(false)
  const [rawDiff, setRawDiff] = useState(false)

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
    setTab('consents'); setLoadingSp(true); clearSelection()
    api.security.servicePrincipals(search, 500)
      .then((r) => {
        setSps(r)
        setStatus((s) => ({ ...s, consents: { ok: true, text: t('security.spsFound', { n: (r || []).length }), at: Date.now() } }))
      })
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setLoadingSp(false))
  }

  const loadRecs = () => {
    setTab('recommendations'); setLoadingRecs(true); setSelRec(null); setRawRecs(false)
    api.security.recommendations()
      .then((r) => {
        const sorted = sortRecs(r || [])
        setRecs(sorted)
        const active = sorted.filter((x) => x.status === 'active' || x.status === 'needsMoreAction').length
        setStatus((s) => ({ ...s, recommendations: { ok: true, text: t('security.recsSummary', { n: sorted.length, active }), at: Date.now() } }))
      })
      .catch((e) => {
        toast('err', errMessage(e))
        setStatus((s) => ({ ...s, recommendations: { ok: false, text: errMessage(e), at: Date.now() } }))
      })
      .finally(() => setLoadingRecs(false))
  }

  const pickRec = (r: GraphObject) => {
    setSelRec(r)
    if (impacted[r.id]) return
    setImpacted((m) => ({ ...m, [r.id]: 'loading' }))
    api.security.recommendationImpacted(r.id)
      .then((list) => setImpacted((m) => ({ ...m, [r.id]: list })))
      .catch((e) => setImpacted((m) => ({ ...m, [r.id]: { error: errMessage(e) } })))
  }

  // Picking B while A's requests are still in flight must not fill B's pane
  // with A's grants: only the newest selection may write.
  const detailReq = useRef(0)
  const loadSpDetail = useCallback((sp: GraphObject) => {
    const ticket = ++detailReq.current
    const fresh = () => detailReq.current === ticket
    api.security.oauthGrants(sp.id).then((r) => { if (fresh()) setGrants(r) }).catch(() => { if (fresh()) setGrants([]) })
    api.security.appRoleAssignments(sp.id).then((r) => { if (fresh()) setAppRoles(r) }).catch(() => { if (fresh()) setAppRoles([]) })
  }, [])

  const pickSp = (sp: GraphObject) => {
    setSelSp(sp); setGrants(null); setAppRoles(null)
    loadSpDetail(sp)
  }

  const clearSelection = () => {
    detailReq.current++ // invalidate anything still in flight
    setSelSp(null); setGrants(null); setAppRoles(null)
  }

  // The selection is restored from the cache on mount, but its grants are not —
  // without this the detail pane spins forever after navigating back.
  useEffect(() => {
    if (selSp && !grants && !appRoles) loadSpDetail(selSp)
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount-time rehydration only
  }, [])

  // --- snapshots ---
  const loadSnaps = useCallback(async () => {
    try {
      const l = await api.snapshot.list()
      setSnaps(l)
      // Default the comparison to the two newest (older → newer).
      setDiffB((cur) => (cur && l.some((m) => m.id === cur) ? cur : l[0]?.id ?? ''))
      setDiffA((cur) => (cur && l.some((m) => m.id === cur) ? cur : l[1]?.id ?? ''))
    } catch (e) {
      toast('err', errMessage(e))
    }
  }, [toast])
  useEffect(() => { loadSnaps() }, [loadSnaps])

  const takeSnapshot = () => {
    setTab('snapshot'); setSnapRaw(null)
    patchJob('snapshot', { running: true, canceled: false, progress: '', log: [], result: null, error: null, startedAt: Date.now() })
    api.snapshot.take(snapName)
      .then(async (meta) => {
        const skipped = (meta.sections || []).filter((s) => s.skipped).length
        const text = t('snapshot.taken', { name: meta.name, n: (meta.sections || []).length, skipped })
        setStatus((s) => ({ ...s, snapshot: { ok: skipped === 0, text, at: Date.now() } }))
        toast(skipped ? 'info' : 'ok', text)
        setSnapName('')
        patchJob('snapshot', { result: meta })
        await loadSnaps()
        setSelSnap(meta)
      })
      .catch((e) => {
        toast('err', errMessage(e))
        patchJob('snapshot', { error: errMessage(e) })
        setStatus((s) => ({ ...s, snapshot: { ok: false, text: errMessage(e), at: Date.now() } }))
      })
      .finally(() => patchJob('snapshot', { running: false, progress: '' }))
  }

  const deleteSnap = (m: SnapshotMeta) => {
    askConfirm(m.name, () => {
      api.snapshot.delete(m.id)
        .then(() => {
          toast('ok', t('snapshot.deleteDone'))
          if (selSnap?.id === m.id) { setSelSnap(null); setSnapRaw(null) }
          return loadSnaps()
        })
        .catch((e) => toast('err', errMessage(e)))
    }, t('snapshot.deleteSnapshot'))
  }

  const pickSnap = (m: SnapshotMeta) => { setSelSnap(m); setSnapRaw(null) }
  const toggleSnapRaw = () => {
    if (!selSnap) return
    if (snapRaw) { setSnapRaw(null); return }
    api.snapshot.get(selSnap.id).then(setSnapRaw).catch((e) => toast('err', errMessage(e)))
  }

  const compare = () => {
    if (!diffA || !diffB) return
    setTab('drift'); setDiffing(true); setRawDiff(false)
    api.snapshot.diff(diffA, diffB)
      .then((d) => {
        setDiff(d)
        setStatus((s) => ({ ...s, drift: { ok: true, text: t('snapshot.summary', { added: d.added, removed: d.removed, changed: d.changed }), at: Date.now() } }))
      })
      .catch((e) => {
        toast('err', errMessage(e))
        setStatus((s) => ({ ...s, drift: { ok: false, text: errMessage(e), at: Date.now() } }))
      })
      .finally(() => setDiffing(false))
  }

  const fmtDate = (v: any) => new Date(v).toLocaleString(i18n.language)
  const snapLabel = (m: SnapshotMeta) => `${m.name} — ${fmtDate(m.takenAt)}`
  const sectionName = (name: string) => t(`snapshot.section.${name}`, { defaultValue: name })

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
    {
      id: 'recommendations', label: t('security.tileRecommendations'), hint: t('security.hintRecommendations'), icon: <Lightbulb size={16} />, variant: 'primary',
      note: <p>{t('security.noteRecommendations')}</p>,
      panel: (
        <TaskForm>
          <Button variant="primary" disabled={loadingRecs} onClick={loadRecs}>
            {loadingRecs ? <Spinner /> : <Lightbulb size={15} />} {t('security.loadRecommendations')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'snapshot', label: t('snapshot.tileTake'), hint: t('snapshot.hintTake'), icon: <Camera size={16} />, variant: 'primary',
      note: <p>{t('snapshot.noteTake')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('snapshot.nameLabel')}>
            <Input value={snapName} placeholder={t('snapshot.namePlaceholder')} onChange={(e) => setSnapName(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && !taking && takeSnapshot()} />
          </Field>
          <Button variant="primary" disabled={taking} onClick={takeSnapshot}>
            {taking ? <Spinner /> : <Camera size={15} />} {t('snapshot.take')}
          </Button>
          {taking && (
            <p className="text-xs text-[var(--accent2)]">
              {progress && progress.section
                ? t('snapshot.taking', { section: sectionName(progress.section), done: progress.done + 1, total: progress.total })
                : t('snapshot.starting')}
            </p>
          )}
          {snaps && snaps.length > 0 && (
            <p className="text-xs text-[var(--text-faint)]">{t('snapshot.existing')}: {snaps.length}</p>
          )}
        </TaskForm>
      ),
    },
    {
      id: 'drift', label: t('snapshot.tileDrift'), hint: t('snapshot.hintDrift'), icon: <GitCompare size={16} />, variant: 'primary',
      note: <p>{t('snapshot.noteDrift')}</p>,
      panel: (
        <TaskForm>
          {(snaps?.length ?? 0) < 2 && <p className="text-xs text-[var(--warn)]">{t('snapshot.needTwo')}</p>}
          <Field label={t('snapshot.older')}>
            <Select value={diffA} onChange={(e) => setDiffA(e.target.value)} className="w-full">
              <option value="">—</option>
              {(snaps || []).map((m) => <option key={m.id} value={m.id}>{snapLabel(m)}</option>)}
            </Select>
          </Field>
          <Field label={t('snapshot.newer')}>
            <Select value={diffB} onChange={(e) => setDiffB(e.target.value)} className="w-full">
              <option value="">—</option>
              {(snaps || []).map((m) => <option key={m.id} value={m.id}>{snapLabel(m)}</option>)}
            </Select>
          </Field>
          <Button variant="primary" disabled={diffing || !diffA || !diffB || diffA === diffB} onClick={compare}>
            {diffing ? <Spinner /> : <GitCompare size={15} />} {t('snapshot.compare')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  const factRows = (rows: [string, React.ReactNode][]) => (
    <dl className="flex flex-col">
      {rows.map(([k, v]) => (
        <div key={k} className="grid grid-cols-[150px_1fr] gap-3 border-b border-[var(--border)]/50 py-1.5 last:border-0">
          <dt className="text-xs uppercase tracking-wide text-[var(--text-faint)]">{k}</dt>
          <dd className="min-w-0 text-sm text-[var(--text)]">{v}</dd>
        </div>
      ))}
    </dl>
  )

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
        {factRows(rows)}
        <details className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-2">
          <summary className="cursor-pointer text-xs text-[var(--text-dim)]">{t('security.rawJson')}</summary>
          <pre className="mt-2 max-h-72 overflow-auto text-xs leading-relaxed text-[var(--text-dim)]">{JSON.stringify(p, null, 2)}</pre>
        </details>
      </div>
    )
  }

  const recDetail = (r: GraphObject) => {
    const imp = impacted[r.id]
    const steps: GraphObject[] = [...(r.actionSteps || [])].sort((a, b) => (a.stepNumber ?? 0) - (b.stepNumber ?? 0))
    const rows: [string, React.ReactNode][] = [
      [t('security.priority'), <Badge kind={priorityBadge[r.priority] || 'neutral'}>{t(`security.priorityLabel.${r.priority}`, { defaultValue: String(r.priority) })}</Badge>],
      [t('security.status'), <Badge kind={statusBadge(r.status)}>{t(`security.statusLabel.${r.status}`, { defaultValue: String(r.status) })}</Badge>],
      [t('security.category'), t(`security.categoryLabel.${r.category}`, { defaultValue: String(r.category ?? '—') })],
      [t('security.impact'), r.impactType === 'tenantLevel' ? t('security.tenantLevel') : String(r.impactType ?? '—')],
    ]
    if (r.maxScore) rows.push([t('security.score'), `${r.currentScore ?? 0} / ${r.maxScore}`])
    rows.push([t('security.insights'), <span className="whitespace-pre-wrap">{r.insights || '—'}</span>])
    rows.push([t('security.benefits'), <span className="whitespace-pre-wrap">{r.benefits || '—'}</span>])
    return (
      <div className="flex flex-col gap-4">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-sm font-semibold">{r.displayName}</span>
          {r.releaseType === 'preview' && <Badge kind="neutral">{t('security.preview')}</Badge>}
        </div>
        {factRows(rows)}
        <div>
          <div className="mb-1 text-sm font-medium">{t('security.actionSteps')}</div>
          {steps.length === 0 && <p className="text-xs text-[var(--text-faint)]">{t('common.empty')}</p>}
          <ol className="flex flex-col gap-1.5">
            {steps.map((s, i) => (
              <li key={i} className="rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs">
                <span className="whitespace-pre-wrap">{s.text}</span>
                {s.actionUrl?.url && (
                  <a href={s.actionUrl.url} target="_blank" rel="noreferrer" className="ml-2 inline-flex items-center gap-1 text-[var(--accent2)] hover:underline">
                    {s.actionUrl.displayName || s.actionUrl.url} <ExternalLink size={11} />
                  </a>
                )}
              </li>
            ))}
          </ol>
        </div>
        <div>
          <div className="mb-1 text-sm font-medium">
            {t('security.impacted')}
            {Array.isArray(imp) && <span className="ml-2 text-xs font-normal text-[var(--text-faint)]">{t('security.impactedCount', { n: imp.length })}</span>}
          </div>
          {(!imp || imp === 'loading') && <Spinner />}
          {imp && typeof imp === 'object' && !Array.isArray(imp) && 'error' in imp && <p className="text-xs text-[var(--danger)]">{imp.error}</p>}
          {Array.isArray(imp) && imp.length === 0 && <p className="text-xs text-[var(--text-faint)]">{r.impactType === 'tenantLevel' ? t('security.tenantLevel') : t('common.empty')}</p>}
          {Array.isArray(imp) && imp.length > 0 && (
            <div className="flex flex-col gap-1">
              {imp.map((x, i) => (
                <div key={x.id ?? i} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs">
                  <span className="min-w-0 flex-1 truncate">{x.displayName || x.subjectId}</span>
                  {x.resourceType && <span className="text-[var(--text-faint)]">{x.resourceType}</span>}
                  {x.status && <Badge kind={statusBadge(x.status)}>{t(`security.statusLabel.${x.status}`, { defaultValue: String(x.status) })}</Badge>}
                  {x.portalUrl && (
                    <a href={x.portalUrl} target="_blank" rel="noreferrer" className="text-[var(--accent2)]" aria-label={String(x.displayName || x.subjectId || t('security.openInPortal'))}><ExternalLink size={12} /></a>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
        <details className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-2">
          <summary className="cursor-pointer text-xs text-[var(--text-dim)]">{t('security.rawView')}</summary>
          <pre className="mt-2 max-h-72 overflow-auto text-xs leading-relaxed text-[var(--text-dim)]">{JSON.stringify(r, null, 2)}</pre>
        </details>
      </div>
    )
  }

  // Facts ⇄ raw switch shared by the panes that offer both.
  const viewToggle = (raw: boolean, setRaw: (v: boolean) => void) => (
    <div className="mb-2 flex gap-1 px-1">
      {([false, true] as const).map((v) => (
        <button key={String(v)} onClick={() => setRaw(v)}
          className={`rounded-md px-2 py-0.5 text-xs font-medium ${raw === v ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]'}`}>
          {v ? t('security.rawView') : t('security.factsView')}
        </button>
      ))}
    </div>
  )

  const snapDetail = (m: SnapshotMeta) => {
    const skipped = (m.sections || []).filter((s) => s.skipped).length
    return (
      <div className="flex flex-col gap-3">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-sm font-semibold">{m.name}</span>
          <span className="text-xs text-[var(--text-faint)]">{fmtDate(m.takenAt)}</span>
          {m.tenant && <span className="text-xs text-[var(--text-faint)]">· {t('snapshot.tenant')}: {m.tenant}</span>}
          {skipped > 0 && <Badge kind="warn">{t('snapshot.skippedN', { n: skipped })}</Badge>}
        </div>
        <div className="text-xs font-mono text-[var(--text-faint)]">{m.id}</div>
        <div>
          <div className="mb-1 text-sm font-medium">{t('snapshot.sections')}</div>
          <div className="flex flex-col gap-1">
            {(m.sections || []).map((s) => (
              <div key={s.name} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs">
                <span className="min-w-0 flex-1">{sectionName(s.name)}</span>
                {s.skipped
                  ? <span className="flex items-center gap-2"><Badge kind="warn">{t('snapshot.skipped')}</Badge><span className="max-w-[320px] truncate text-[var(--text-faint)]" title={s.error ? errMessage(s.error) : ''}>{s.error ? errMessage(s.error) : ''}</span></span>
                  : <span className="text-[var(--text-dim)]">{t('snapshot.objectsN', { n: s.count })}</span>}
              </div>
            ))}
          </div>
        </div>
        <div className="flex gap-2">
          <Button variant="ghost" onClick={toggleSnapRaw}>{snapRaw ? t('snapshot.hideRaw') : t('snapshot.showRaw')}</Button>
          <Button variant="danger" onClick={() => deleteSnap(m)}><Trash2 size={14} /> {t('snapshot.deleteSnapshot')}</Button>
        </div>
        {snapRaw && (
          <div className="h-[420px] overflow-hidden rounded-lg border border-[var(--border)]">
            <ResultView data={snapRaw} />
          </div>
        )}
      </div>
    )
  }

  const objectRow = (o: ObjectChange, kind: 'added' | 'removed') => (
    <details key={`${kind}:${o.key}`} className="rounded-lg border border-[var(--border)] bg-[var(--bg)]">
      <summary className="flex cursor-pointer items-center gap-2 px-3 py-1.5 text-xs">
        <Badge kind={kind === 'added' ? 'ok' : 'danger'}>{kind === 'added' ? '+' : '−'} {t(`snapshot.${kind}`)}</Badge>
        <span className="min-w-0 flex-1 truncate">{o.label}</span>
        {o.key !== o.label && <span className="truncate font-mono text-[var(--text-faint)]">{o.key}</span>}
      </summary>
      <pre className="max-h-64 overflow-auto border-t border-[var(--border)] p-2 text-[11px] leading-snug text-[var(--text-dim)]">{JSON.stringify(o.object ?? {}, null, 2)}</pre>
    </details>
  )

  const changedRow = (o: ObjectChange) => (
    <details key={`changed:${o.key}`} open className="rounded-lg border border-[var(--border)] bg-[var(--bg)]">
      <summary className="flex cursor-pointer items-center gap-2 px-3 py-1.5 text-xs">
        <Badge kind="warn">~ {t('snapshot.changed')}</Badge>
        <span className="min-w-0 flex-1 truncate">{o.label}</span>
        <span className="text-[var(--text-faint)]">{o.changes?.length ?? 0}</span>
      </summary>
      <table className="w-full border-t border-[var(--border)] text-xs">
        <thead>
          <tr className="text-left text-[10px] uppercase tracking-wide text-[var(--text-faint)]">
            <th className="px-3 py-1 font-medium">{t('snapshot.field')}</th>
            <th className="px-3 py-1 font-medium">{t('snapshot.before')}</th>
            <th className="px-3 py-1 font-medium">{t('snapshot.after')}</th>
          </tr>
        </thead>
        <tbody>
          {(o.changes || []).map((c, i) => (
            <tr key={i} className="border-t border-[var(--border)]/50 align-top">
              <td className="px-3 py-1 font-mono text-[11px] text-[var(--accent2)]">{c.path}</td>
              <td className="px-3 py-1 text-[var(--danger)]"><DiffValue v={c.before} /></td>
              <td className="px-3 py-1 text-[var(--ok)]"><DiffValue v={c.after} /></td>
            </tr>
          ))}
        </tbody>
      </table>
    </details>
  )

  const sectionRow = (s: SectionDiff) => {
    const n = (s.added?.length ?? 0) + (s.removed?.length ?? 0) + (s.changed?.length ?? 0)
    return (
      <details key={s.name} open={n > 0} className="rounded-xl border border-[var(--border)] bg-[var(--bg-elev-2)]/40">
        <summary className="flex cursor-pointer flex-wrap items-center gap-2 px-3 py-2 text-sm">
          <span className="min-w-0 flex-1 font-medium">{sectionName(s.name)}</span>
          {s.skipped ? <Badge kind="neutral">{t('snapshot.notCompared')}</Badge> : (
            <>
              {(s.added?.length ?? 0) > 0 && <Badge kind="ok">+{s.added.length}</Badge>}
              {(s.removed?.length ?? 0) > 0 && <Badge kind="danger">−{s.removed.length}</Badge>}
              {(s.changed?.length ?? 0) > 0 && <Badge kind="warn">~{s.changed.length}</Badge>}
              <span className="text-xs text-[var(--text-faint)]">{t('snapshot.unchangedN', { n: s.unchanged })}</span>
            </>
          )}
        </summary>
        <div className="flex flex-col gap-1.5 border-t border-[var(--border)] p-2">
          {s.skipped && <p className="text-xs text-[var(--text-faint)]">{s.note ? errMessage(s.note) : ''}</p>}
          {!s.skipped && n === 0 && <p className="text-xs text-[var(--text-faint)]">{t('snapshot.noChangesSection')}</p>}
          {(s.added || []).map((o) => objectRow(o, 'added'))}
          {(s.removed || []).map((o) => objectRow(o, 'removed'))}
          {(s.changed || []).map(changedRow)}
        </div>
      </details>
    )
  }

  const driftPane = (d: SnapshotDiff) => {
    const total = d.added + d.removed + d.changed
    return (
      <div className="flex h-full flex-col">
        <div className="flex flex-wrap items-center gap-3 border-b border-[var(--border)] px-4 py-2 text-xs">
          <span className="text-[var(--text-dim)]">A: <span className="text-[var(--text)]">{snapLabel(d.a)}</span></span>
          <span className="text-[var(--text-faint)]">→</span>
          <span className="text-[var(--text-dim)]">B: <span className="text-[var(--text)]">{snapLabel(d.b)}</span></span>
          <span className="ml-auto font-medium">{total === 0 ? t('snapshot.noChanges') : t('snapshot.summary', { added: d.added, removed: d.removed, changed: d.changed })}</span>
        </div>
        <div className="min-h-0 flex-1 overflow-auto p-3">
          {viewToggle(rawDiff, setRawDiff)}
          {rawDiff
            ? <div className="h-[calc(100%-2rem)] overflow-hidden rounded-lg border border-[var(--border)]"><ResultView data={d as any} /></div>
            : <div className="flex flex-col gap-2">{(d.sections || []).map(sectionRow)}</div>}
        </div>
      </div>
    )
  }

  const listPane = (
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
      {tab === 'recommendations' && viewToggle(rawRecs, setRawRecs)}
      {tab === 'recommendations' && !rawRecs && (recs || []).map((r) => (
        <button key={r.id} onClick={() => pickRec(r)}
          className={`mb-1 flex w-full flex-col gap-1 rounded-lg border px-3 py-2 text-left text-sm ${selRec?.id === r.id ? 'border-[var(--accent)] bg-[var(--accent)]/10' : 'border-transparent hover:bg-[var(--bg-elev-2)]'}`}>
          <span className="w-full truncate">{r.displayName}</span>
          <span className="flex flex-wrap gap-1">
            <Badge kind={priorityBadge[r.priority] || 'neutral'}>{t(`security.priorityLabel.${r.priority}`, { defaultValue: String(r.priority) })}</Badge>
            <Badge kind={statusBadge(r.status)}>{t(`security.statusLabel.${r.status}`, { defaultValue: String(r.status) })}</Badge>
          </span>
        </button>
      ))}
      {tab === 'snapshot' && (snaps || []).map((m) => {
        const skipped = (m.sections || []).filter((s) => s.skipped).length
        return (
          <button key={m.id} onClick={() => pickSnap(m)}
            className={`mb-1 flex w-full flex-col gap-0.5 rounded-lg border px-3 py-2 text-left text-sm ${selSnap?.id === m.id ? 'border-[var(--accent)] bg-[var(--accent)]/10' : 'border-transparent hover:bg-[var(--bg-elev-2)]'}`}>
            <span className="flex w-full items-center gap-2">
              <span className="min-w-0 flex-1 truncate">{m.name}</span>
              {skipped > 0 && <Badge kind="warn">{t('snapshot.skippedN', { n: skipped })}</Badge>}
            </span>
            <span className="text-xs text-[var(--text-faint)]">{fmtDate(m.takenAt)}</span>
          </button>
        )
      })}
      {tab === 'ca' && policies && policies.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
      {tab === 'consents' && sps && sps.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
      {tab === 'recommendations' && recs && recs.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
      {tab === 'snapshot' && snaps && snaps.length === 0 && <p className="p-2 text-sm text-[var(--text-faint)]">{t('snapshot.none')}</p>}
    </div>
  )

  const detailPane = (
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

      {tab === 'recommendations' && !selRec && <p className="text-sm text-[var(--text-faint)]">{t('security.pickRec')}</p>}
      {tab === 'recommendations' && selRec && recDetail(selRec)}

      {tab === 'snapshot' && !selSnap && <p className="text-sm text-[var(--text-faint)]">{t('snapshot.pickSnapshot')}</p>}
      {tab === 'snapshot' && selSnap && snapDetail(selSnap)}
    </div>
  )

  const resultPane = tab === 'drift'
    ? (diff ? driftPane(diff) : <p className="p-4 text-sm text-[var(--text-faint)]">{t('snapshot.needTwo')}</p>)
    : tab === 'recommendations' && rawRecs
      ? (
        <div className="grid h-full grid-cols-1 lg:grid-cols-[minmax(240px,320px)_1fr]">
          {listPane}
          <div className="min-h-0 overflow-hidden"><ResultView data={selRec ? [selRec] : recs} /></div>
        </div>
      )
      : (
        <div className="grid h-full grid-cols-1 lg:grid-cols-[minmax(240px,320px)_1fr]">
          {listPane}
          {detailPane}
        </div>
      )

  const busy = loadingCa || loadingSp || loadingRecs || taking || diffing
  const busyLabel = loadingCa ? t('security.loadPolicies')
    : loadingSp ? t('common.search')
      : loadingRecs ? t('security.loadRecommendations')
        : taking ? (progress?.section ? t('snapshot.taking', { section: sectionName(progress.section), done: progress.done + 1, total: progress.total }) : t('snapshot.starting'))
          : t('snapshot.comparing')

  const hasResult = (tab === 'ca' && !!policies) || (tab === 'consents' && !!sps) || (tab === 'recommendations' && !!recs)
    || (tab === 'snapshot' && !!snaps) || (tab === 'drift' && !!diff) || busy

  const clearResult = () => {
    switch (tab) {
      case 'ca': setPolicies(null); setSelPolicy(null); setCache('security.policies', null); break
      case 'consents': setSps(null); clearSelection(); break
      case 'recommendations': setRecs(null); setSelRec(null); setImpacted({}); break
      case 'snapshot': setSelSnap(null); setSnapRaw(null); setTab('ca'); break
      case 'drift': setDiff(null); break
    }
  }

  return (
    <>
      <TaskPage
        pageId="security"
        title={t('security.title')}
        subtitle={t('security.subtitle')}
        actions={actions}
        status={status}
        busy={busy}
        busyLabel={busyLabel}
        hasResult={hasResult}
        onClearResult={clearResult}
        result={resultPane}
      />
      {confirmElement}
    </>
  )
}
