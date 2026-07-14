import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { RefreshCw, CheckCircle2, XCircle } from 'lucide-react'
import { Page } from '../components/Layout'
import { Button, Card } from '../components/ui'
import { api } from '../lib/api'

export function ActivityPage() {
  const { t } = useTranslation()
  const [entries, setEntries] = useState<any[]>([])

  const load = () => api.audit.activity(200).then((e) => setEntries((e || []).reverse())).catch(() => {})
  useEffect(() => { load() }, [])

  return (
    <Page title={t('nav.activity')} subtitle="Local action log (destructive & write operations)">
      <Card title={t('nav.activity')} actions={<Button variant="ghost" onClick={load} className="!px-2 !py-1"><RefreshCw size={14} /></Button>}>
        {entries.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
        <div className="flex flex-col gap-1">
          {entries.map((e, i) => (
            <div key={i} className="flex items-center gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm">
              {e.ok ? <CheckCircle2 size={15} className="text-[var(--ok)]" /> : <XCircle size={15} className="text-[var(--danger)]" />}
              <span className="font-mono text-xs text-[var(--accent)]">{e.action}</span>
              <span className="truncate text-[var(--text-dim)]">{e.target}</span>
              {e.detail && <span className="truncate text-xs text-[var(--text-faint)]">{e.detail}</span>}
              <span className="ml-auto shrink-0 text-xs text-[var(--text-faint)]">
                {new Date(e.time).toLocaleString()}
              </span>
            </div>
          ))}
        </div>
      </Card>
    </Page>
  )
}
