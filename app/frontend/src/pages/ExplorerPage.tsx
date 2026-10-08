import { Fragment, useEffect, useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Users, Boxes, MessagesSquare, ChevronRight } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Input, Spinner, ErrorNote } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import { useCatalog, entryLabel, prefillAction } from '../components/CatalogAction'

type Kind = 'users' | 'groups' | 'teams'

const KINDS: { id: Kind; icon: JSX.Element; field: 'user' | 'group' }[] = [
  { id: 'users', icon: <Users size={15} />, field: 'user' },
  { id: 'groups', icon: <Boxes size={15} />, field: 'group' },
  { id: 'teams', icon: <MessagesSquare size={15} />, field: 'group' },
]

// Object navigator: browse the tenant's objects, then open any catalog action
// that takes the chosen one, with the object already filled in.
export function ExplorerPage() {
  const { t } = useTranslation()
  const { connected, goTo, cache, setCache } = useStore()
  const catalog = useCatalog()
  const [kind, setKind] = useState<Kind>('users')
  const [q, setQ] = useState('')
  const [items, setItems] = useState<GraphObject[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [picked, setPicked] = useState<GraphObject | null>(null)

  useEffect(() => {
    if (!connected) return
    let alive = true
    setItems(null); setError(null)
    const timer = setTimeout(async () => {
      try {
        let list: GraphObject[]
        if (kind === 'users') list = await api.users.list(q, 50)
        else if (kind === 'groups') list = await api.groups.list(q, 50)
        else {
          const needle = q.trim().toLowerCase()
          // All teams come in one call: fetch once, filter as the operator types.
          let all = cache['explorer.teams'] as GraphObject[] | undefined
          if (!all) { all = await api.teams.all(); setCache('explorer.teams', all) }
          list = all.filter((x: GraphObject) => !needle || (x.displayName || '').toLowerCase().includes(needle)).slice(0, 50)
        }
        if (alive) setItems(list)
      } catch (e) { if (alive) setError(errMessage(e)) }
    }, 300)
    return () => { alive = false; clearTimeout(timer) }
  }, [kind, q, connected]) // eslint-disable-line react-hooks/exhaustive-deps -- cache is read, not watched

  const field = KINDS.find((k) => k.id === kind)!.field
  // Write and read actions whose inputs take this kind of object.
  const actions = useMemo(() => (catalog || [])
    .filter((e) => e.available && e.fields.some((f) => f.kind === field))
    .sort((a, b) => entryLabel(a, t).localeCompare(entryLabel(b, t))), [catalog, field, t])

  const open = (entryId: string, page: string) => {
    if (!picked) return
    const f = catalog?.find((e) => e.id === entryId)?.fields.find((x) => x.kind === field)
    if (!f) return
    // Users go by UPN (what the pickers fill in), groups by id.
    prefillAction(entryId, { [f.name]: field === 'user' ? (picked.userPrincipalName || picked.id) : picked.id })
    goTo(page, entryId)
  }

  return (
    <Page title={t('nav.explorer')}>
      {!connected && <p className="text-sm text-[var(--text-faint)]">{t('common.notConnected')}</p>}
      {connected && (
        <div className="grid min-h-0 grid-cols-1 gap-4 lg:grid-cols-[minmax(0,1fr)_minmax(0,1fr)]">
          <Card title={t('explorer.objects')}>
            <div className="mb-2 flex gap-1">
              {KINDS.map((k) => (
                <button key={k.id} onClick={() => { setKind(k.id); setPicked(null) }}
                  className={`flex items-center gap-1.5 rounded-md px-2.5 py-1 text-xs ${kind === k.id ? 'bg-[var(--accent)]/15 text-[var(--accent)]' : 'text-[var(--text-dim)] hover:bg-[var(--surface)]'}`}>
                  {k.icon} {t(`explorer.kinds.${k.id}`)}
                </button>
              ))}
            </div>
            <Input value={q} onChange={(e) => setQ(e.target.value)} placeholder={t('explorer.search')} />
            {error && <ErrorNote>{error}</ErrorNote>}
            {!items && !error && <div className="py-4"><Spinner /></div>}
            <div className="mt-2 flex max-h-[60vh] flex-col overflow-auto">
              {items?.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
              {items?.map((o) => (
                <button key={o.id} onClick={() => setPicked(o)}
                  className={`flex items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm ${picked?.id === o.id ? 'bg-[var(--accent)]/10' : 'hover:bg-[var(--surface)]'}`}>
                  <div className="min-w-0 flex-1">
                    <div className="truncate">{o.displayName || o.id}</div>
                    <div className="truncate text-xs text-[var(--text-faint)]">{o.userPrincipalName || o.mail || o.id}</div>
                  </div>
                  <ChevronRight size={14} className="text-[var(--text-faint)]" />
                </button>
              ))}
            </div>
          </Card>
          <Card title={picked ? (picked.displayName || picked.id) : t('explorer.pick')}>
            {picked && (
              <>
                <dl className="mb-3 grid grid-cols-[auto_1fr] gap-x-3 gap-y-1 text-xs">
                  {(['userPrincipalName', 'mail', 'jobTitle', 'department', 'id'] as const).filter((k) => picked[k]).map((k) => (
                    <Fragment key={k}><dt className="text-[var(--text-faint)]">{t(`explorer.props.${k}`)}</dt><dd className="truncate">{String(picked[k])}</dd></Fragment>
                  ))}
                </dl>
                <p className="mb-1 text-xs font-medium text-[var(--text-dim)]">{t('explorer.actions')}</p>
                {!catalog && <Spinner />}
                <div className="flex flex-col gap-1">
                  {actions.map((e) => (
                    <button key={e.id} onClick={() => open(e.id, e.page)}
                      className="flex items-center justify-between rounded-md border border-[var(--border)] px-2.5 py-1.5 text-left text-sm hover:bg-[var(--surface)]">
                      {entryLabel(e, t)}
                      <span className="text-xs text-[var(--text-faint)]">{t(`nav.${e.page}`, { defaultValue: e.page })}</span>
                    </button>
                  ))}
                </div>
              </>
            )}
          </Card>
        </div>
      )}
    </Page>
  )
}
