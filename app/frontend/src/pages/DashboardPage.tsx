import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Users, Boxes, Globe, KeyRound, RefreshCw } from 'lucide-react'
import type { ReactNode } from 'react'
import { Page } from '../components/Layout'
import { Card, Button, Spinner, ErrorNote } from '../components/ui'
import { api, errMessage } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

export function DashboardPage() {
  const { t } = useTranslation()
  const [data, setData] = useState<services.DashboardSummary | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const load = async () => {
    setLoading(true); setError(null)
    try { setData(await api.dashboard.summary()) } catch (e) { setError(errMessage(e)) } finally { setLoading(false) }
  }
  useEffect(() => { load() }, [])

  const topLicenses = (data?.licenses || [])
    .filter((l) => l.total > 0)
    .sort((a, b) => b.consumed - a.consumed)
    .slice(0, 8)

  return (
    <Page title={data?.orgName || t('dashboard.title')}>
      <div className="mb-4 flex justify-end">
        <Button variant="ghost" onClick={load} disabled={loading}>
          {loading ? <Spinner /> : <RefreshCw size={15} />} {t('dashboard.refresh')}
        </Button>
      </div>

      {error && <ErrorNote>{error}</ErrorNote>}

      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <Stat icon={<Users size={20} />} label={t('dashboard.users')} value={data?.users} loading={loading} />
        <Stat icon={<Boxes size={20} />} label={t('dashboard.groups')} value={data?.groups} loading={loading} />
        <Stat icon={<Globe size={20} />} label={t('dashboard.domains')} value={data?.domains} loading={loading} />
        <Stat
          icon={<KeyRound size={20} />}
          label={t('dashboard.licenses')}
          value={data ? `${data.licensesUsed}/${data.licensesTotal}` : undefined}
          loading={loading}
        />
      </div>

      {topLicenses.length > 0 && (
        <Card title={t('dashboard.topLicenses')} className="mt-4">
          <div className="flex flex-col gap-3">
            {topLicenses.map((l) => {
              const pct = l.total > 0 ? Math.round((l.consumed / l.total) * 100) : 0
              return (
                <div key={l.skuPartNumber}>
                  <div className="mb-1 flex justify-between text-sm">
                    <span className="font-mono text-[var(--text)]">{l.skuPartNumber}</span>
                    <span className="text-[var(--text-dim)]">{t('dashboard.licensesUsed', { used: l.consumed, total: l.total })}</span>
                  </div>
                  <div className="h-2 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                    <div
                      className="h-full rounded transition-all"
                      style={{ width: `${pct}%`, background: pct >= 90 ? 'var(--danger)' : pct >= 70 ? 'var(--warn)' : 'var(--accent)' }}
                    />
                  </div>
                </div>
              )
            })}
          </div>
        </Card>
      )}
    </Page>
  )
}

function Stat({ icon, label, value, loading }: { icon: ReactNode; label: string; value?: number | string; loading?: boolean }) {
  return (
    <div className="rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] p-4">
      <div className="flex items-center gap-2 text-[var(--accent)]">{icon}</div>
      <div className="mt-2 text-2xl font-semibold text-[var(--text)]">
        {loading ? <Spinner /> : (value ?? '—')}
      </div>
      <div className="text-sm text-[var(--text-dim)]">{label}</div>
    </div>
  )
}
