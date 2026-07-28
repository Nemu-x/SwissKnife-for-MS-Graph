import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, AlarmClock, KeyRound, Copy, List } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Badge, Spinner } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

export function AppsPage() {
  const { t } = useTranslation()
  const { readOnly, toast, cache, setCache } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [days, setDays] = useState(30)
  // Cache-backed: the tenant-wide expiring scan is slow, and a fresh secret is a
  // one-time value — neither should be lost by navigating away.
  const [expiring, setExpiringLocal] = useState<services.ExpiringCredential[] | null>(() => cache['apps.expiring'] ?? null)
  const setExpiring = (v: services.ExpiringCredential[] | null) => { setExpiringLocal(v); setCache('apps.expiring', v) }
  const [error, setError] = useState<string | null>(null)
  // The expiring-credentials scan walks every registration: it needs its own
  // busy flag, it is not a useAsync call.
  const [scanning, setScanning] = useState(false)
  const [rotateId, setRotateId] = useState('')
  const [rotateName, setRotateName] = useState('')
  const [rotateMonths, setRotateMonths] = useState(6)
  const [newSecret, setNewSecretLocal] = useState<GraphObject | null>(() => cache['apps.newSecret'] ?? null)
  const setNewSecret = (v: GraphObject | null) => { setNewSecretLocal(v); setCache('apps.newSecret', v) }
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  const listApps = () => { setExpiring(null); res.run(() => api.apps.list(search, 0)) }
  const loadExpiring = async () => {
    setError(null)
    setScanning(true)
    try { setExpiring(await api.apps.expiring(days)) } catch (e) { setError(errMessage(e)) }
    finally { setScanning(false) }
  }
  const rotate = async () => {
    try {
      setNewSecret(await api.apps.addSecret(rotateId, rotateName, rotateMonths))
      setStatus((s) => ({ ...s, rotate: { ok: true, text: t('apps.rotate'), at: Date.now() } }))
    } catch (e) {
      const m = errMessage(e)
      setStatus((s) => ({ ...s, rotate: { ok: false, text: m, at: Date.now() } }))
      toast('err', m)
    }
  }

  const actions: TaskAction[] = [
    { id: 'list', label: t('apps.tileList'), hint: t('apps.hintList'), icon: <List size={16} />, variant: 'primary', onClick: listApps },
    {
      id: 'expiring', label: t('apps.tileExpiring'), hint: t('apps.hintExpiring'), icon: <AlarmClock size={16} />, variant: 'primary',
      note: <p>{t('apps.noteExpiring')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('apps.withinDays')}><Input type="number" value={days} onChange={(e) => setDays(Number(e.target.value) || 30)} /></Field>
          <Button variant="primary" disabled={scanning} onClick={loadExpiring}>
            {scanning ? <Spinner /> : <AlarmClock size={15} />} {scanning ? t('apps.scanning') : t('common.run')}
          </Button>
          {expiring !== null && (
            <Button variant="ghost" onClick={() => setExpiring(null)}>{t('apps.backToList')}</Button>
          )}
        </TaskForm>
      ),
    },
    {
      id: 'rotate', label: t('apps.tileRotate'), hint: t('apps.hintRotate'), icon: <KeyRound size={16} />, write: true,
      note: <p className="text-[var(--warn)]">{t('apps.newSecretHint')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('apps.objectId')} hint={t('apps.objectIdHint')}>
            <Input value={rotateId} placeholder="object id" onChange={(e) => setRotateId(e.target.value)} />
          </Field>
          <Field label={t('apps.secretName')}>
            <Input value={rotateName} placeholder="rotation-2026" onChange={(e) => setRotateName(e.target.value)} />
          </Field>
          <Field label={t('apps.secretMonths')}>
            <Input type="number" value={rotateMonths} onChange={(e) => setRotateMonths(Math.min(24, Math.max(1, Number(e.target.value) || 6)))} />
          </Field>
          <Button variant="primary" disabled={readOnly || !rotateId} onClick={rotate}>
            <KeyRound size={15} /> {t('apps.rotate')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {newSecret && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60" onClick={() => setNewSecret(null)}>
          <div className="w-[460px] rounded-2xl border border-[var(--border)] bg-[var(--bg-elev)] p-5" onClick={(e) => e.stopPropagation()}>
            <div className="mb-2 text-sm font-medium">{t('apps.newSecretTitle')}</div>
            <div className="break-all rounded-lg bg-[var(--bg)] p-3 text-center font-mono text-sm">{String(newSecret.secretText || '')}</div>
            <div className="mt-2 text-center text-xs text-[var(--warn)]">{t('apps.newSecretHint')}</div>
            <div className="mt-3 flex gap-2">
              <Button variant="primary" className="flex-1"
                onClick={() => { navigator.clipboard.writeText(String(newSecret.secretText || '')); toast('ok', t('apps.copied')) }}>
                <Copy size={15} /> {t('apps.copy')}
              </Button>
              <Button className="flex-1" onClick={() => setNewSecret(null)}>{t('apps.close')}</Button>
            </div>
          </div>
        </div>
      )}
      <TaskPage
        pageId="apps"
        title={t('apps.title')}
        subtitle={t('apps.subtitle')}
        search={{ value: search, onChange: setSearch, onSubmit: listApps, placeholder: t('common.search') }}
        actions={actions}
        status={status}
        busy={res.loading || scanning}
        busyLabel={scanning ? t('apps.scanning') : undefined}
        onClearResult={() => { res.reset(); setExpiring(null); setError(null) }}
        hasResult={expiring !== null || !!res.data || res.loading || !!res.error}
        result={
          expiring !== null
            ? <ExpiringList items={expiring} error={error} />
            : <ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setRotateId} />
        }
      />
    </>
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
            {[t('apps.app'), t('apps.kind'), t('apps.secretName'), t('apps.expires'), t('apps.daysLeft')].map((h) => (
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
                  {c.daysLeft < 0 ? t('apps.expiredDays', { n: -c.daysLeft }) : t('apps.daysShort', { n: c.daysLeft })}
                </Badge>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
