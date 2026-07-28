import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Users, Boxes, Globe, KeyRound, RefreshCw } from 'lucide-react'
import type { ReactNode } from 'react'
import { Page } from '../components/Layout'
import { Card, Button, Spinner, ErrorNote } from '../components/ui'
import { api, errMessage } from '../lib/api'
import { useStore } from '../lib/store'
import { skuFriendly, isFreeOrTrial } from '../lib/skuNames'
import type { services } from '../../wailsjs/go/models'

export function DashboardPage() {
  const { t } = useTranslation()
  const { goTo } = useStore()
  const [data, setData] = useState<services.DashboardSummary | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const load = async () => {
    setLoading(true); setError(null)
    try { setData(await api.dashboard.summary()) } catch (e) { setError(errMessage(e)) } finally { setLoading(false) }
  }
  useEffect(() => { load() }, [])

  const all = data?.licenses || []
  const paid = all.filter((l) => l.total > 0 && !isFreeOrTrial(l.skuPartNumber, l.total)).sort((a, b) => b.consumed - a.consumed)
  const free = all.filter((l) => l.consumed > 0 && isFreeOrTrial(l.skuPartNumber, l.total)).sort((a, b) => b.consumed - a.consumed)

  return (
    <Page title={data?.orgName || t('dashboard.title')} subtitle={t('dashboard.subtitle')}>
      <div className="mb-4 flex justify-end">
        <Button variant="ghost" onClick={load} disabled={loading}>
          {loading ? <Spinner /> : <RefreshCw size={15} />} {t('dashboard.refresh')}
        </Button>
      </div>

      {error && <ErrorNote>{error}</ErrorNote>}

      {/* Every number opens the list behind it — a dashboard that cannot be
          drilled into is a dead end. */}
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <Stat icon={<Users size={18} />} label={t('dashboard.users')} value={data?.users} loading={loading}
          onClick={() => goTo('users', 'list')} />
        <Stat icon={<Boxes size={18} />} label={t('dashboard.groups')} value={data?.groups} loading={loading}
          onClick={() => goTo('groups', 'list')} />
        <Stat icon={<Globe size={18} />} label={t('dashboard.domains')} value={data?.domains} loading={loading}
          onClick={() => goTo('connect')} />
        <Stat
          icon={<KeyRound size={18} />}
          label={t('dashboard.licensesInUse')}
          value={data?.licensesUsed}
          suffix={t('dashboard.assigned')}
          loading={loading}
          onClick={() => goTo('licensing', 'skus')}
        />
      </div>

      <div className="mt-5 grid grid-cols-1 gap-4 lg:grid-cols-[1.5fr_1fr]">
        <Card title={t('dashboard.paidLicenses')}>
          {paid.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          <div className="flex flex-col gap-4">
            {paid.map((l) => {
              const pct = l.total > 0 ? Math.round((l.consumed / l.total) * 100) : 0
              const left = l.total - l.consumed
              const color = left <= 0 ? 'var(--danger)' : left <= Math.max(1, l.total * 0.1) ? 'var(--warn)' : 'var(--accent)'
              return (
                <div key={l.skuPartNumber}>
                  <div className="mb-1.5 flex items-baseline justify-between gap-3">
                    <span className="truncate text-sm font-medium text-[var(--text)]" title={l.skuPartNumber}>
                      {skuFriendly(l.skuPartNumber)}
                    </span>
                    <span className="shrink-0 text-xs tabular-nums text-[var(--text-dim)]">
                      {t('dashboard.seatsLeft', { used: l.consumed, total: l.total, left })}
                    </span>
                  </div>
                  <div className="h-2 overflow-hidden rounded-full bg-[var(--bg-elev-2)]">
                    <div className="h-full rounded-full transition-all" style={{ width: `${pct}%`, background: color }} />
                  </div>
                </div>
              )
            })}
          </div>
        </Card>

        <Card title={t('dashboard.freeTrial')}>
          {free.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          <div className="flex flex-wrap gap-2">
            {free.map((l) => (
              <div
                key={l.skuPartNumber}
                className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg-elev-2)] px-3 py-1.5"
                title={l.skuPartNumber}
              >
                <span className="text-xs text-[var(--text-dim)]">{skuFriendly(l.skuPartNumber)}</span>
                <span className="text-sm font-semibold tabular-nums text-[var(--text)]">{l.consumed}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Page>
  )
}

function Stat({ icon, label, value, suffix, loading, onClick }: {
  icon: ReactNode; label: string; value?: number | string; suffix?: string; loading?: boolean; onClick?: () => void
}) {
  const Tag = onClick ? 'button' : 'div'
  return (
    <Tag
      onClick={onClick}
      className={`rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] p-4 text-left
        ${onClick ? 'transition-colors hover:border-[var(--accent)] hover:bg-[var(--bg-elev-2)]' : ''}`}
    >
      <div className="mb-3 flex h-9 w-9 items-center justify-center rounded-[10px] text-[var(--accent)]"
        style={{ background: 'color-mix(in srgb, var(--accent) 14%, transparent)' }}>
        {icon}
      </div>
      <div className="text-2xl font-bold tabular-nums tracking-tight text-[var(--text)]">
        {loading ? <Spinner /> : (value ?? '—')}
        {!loading && suffix && value != null && <span className="ml-1 text-sm font-medium text-[var(--text-faint)]">{suffix}</span>}
      </div>
      <div className="mt-0.5 text-sm text-[var(--text-dim)]">{label}</div>
    </Tag>
  )
}
