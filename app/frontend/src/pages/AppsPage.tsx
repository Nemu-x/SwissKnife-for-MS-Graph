import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, AlarmClock } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Badge } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { api, errMessage, type GraphObject } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

export function AppsPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [days, setDays] = useState(30)
  const [expiring, setExpiring] = useState<services.ExpiringCredential[] | null>(null)
  const [error, setError] = useState<string | null>(null)

  const listApps = () => { setExpiring(null); res.run(() => api.apps.list(search, 0)) }
  const loadExpiring = async () => {
    setError(null)
    try { setExpiring(await api.apps.expiring(days)) } catch (e) { setError(errMessage(e)) }
  }

  const actions: Action[] = [
    { id: 'search', label: t('common.search'), icon: <Search size={15} />, variant: 'primary', onClick: listApps },
    {
      id: 'expiring', label: t('apps.expiring'), icon: <AlarmClock size={15} />,
      panel: (
        <DrawerForm>
          <Field label={t('apps.withinDays')}><Input type="number" value={days} onChange={(e) => setDays(Number(e.target.value) || 30)} /></Field>
          <Button variant="primary" onClick={loadExpiring}><AlarmClock size={15} /> {t('common.run')}</Button>
        </DrawerForm>
      ),
    },
  ]

  return (
    <ActionPage
      title={t('apps.title')}
      search={{ value: search, onChange: setSearch, onSubmit: listApps, placeholder: t('common.search') }}
      actions={actions}
      result={
        expiring !== null
          ? <ExpiringList items={expiring} error={error} />
          : <ResultView data={res.data} loading={res.loading} error={res.error} />
      }
    />
  )
}

function ExpiringList({ items, error }: { items: services.ExpiringCredential[]; error: string | null }) {
  const { t } = useTranslation()
  if (error) return <div className="p-4 text-sm text-[var(--danger)]">{error}</div>
  if (items.length === 0) return <div className="flex h-full items-center justify-center text-sm text-[var(--text-faint)]">{t('apps.none')}</div>
  return (
    <div className="h-full overflow-auto">
      <table className="w-full border-collapse text-sm">
        <thead className="sticky top-0 bg-[var(--bg-elev-2)]">
          <tr>
            {[t('apps.app'), t('apps.kind'), 'Name', t('apps.expires'), t('apps.daysLeft')].map((h) => (
              <th key={h} className="border-b border-[var(--border)] px-3 py-2 text-left font-semibold text-[var(--text-dim)]">{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {items.map((c, i) => (
            <tr key={i} className="hover:bg-[var(--bg-elev-2)]/50">
              <td className="border-b border-[var(--border)] px-3 py-2">{c.appName}</td>
              <td className="border-b border-[var(--border)] px-3 py-2"><Badge kind={c.kind === 'certificate' ? 'neutral' : 'warn'}>{c.kind}</Badge></td>
              <td className="border-b border-[var(--border)] px-3 py-2 text-[var(--text-dim)]">{c.displayName || '—'}</td>
              <td className="border-b border-[var(--border)] px-3 py-2 tabular-nums">{c.expires}</td>
              <td className="border-b border-[var(--border)] px-3 py-2">
                <Badge kind={c.daysLeft < 0 ? 'danger' : c.daysLeft <= 7 ? 'warn' : 'neutral'}>
                  {c.daysLeft < 0 ? `expired ${-c.daysLeft}d` : `${c.daysLeft}d`}
                </Badge>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
