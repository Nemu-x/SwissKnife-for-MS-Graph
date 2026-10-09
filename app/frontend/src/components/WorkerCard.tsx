import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Link2, Unlink, RefreshCw, Server } from 'lucide-react'
import { Card, Button, Field, Input, Badge } from './ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'

type Worker = {
  addr: string; fingerprint: string; name: string; pairedAt: string; families?: string[]
  unavailable?: string[]; version?: string; lastSeen?: string; online: boolean; lastError?: string
}

const familyName = (f: string) => (f === 'teams' ? 'Teams PS' : 'Exchange PS')

// A Go zero time (never seen) comes back as year 1.
const when = (iso?: string) => {
  if (!iso) return ''
  const d = new Date(iso)
  return d.getFullYear() < 2000 ? '' : d.toLocaleString()
}

// A paired Windows machine that runs PowerShell for this app when this one
// cannot (no PowerShell, a module missing). Pairing pins both certificates.
export function WorkerCard() {
  const { t } = useTranslation()
  const { toast, setCache } = useStore()
  const [w, setW] = useState<Worker | null>(null)
  const [me, setMe] = useState('')
  const [addr, setAddr] = useState('')
  const [code, setCode] = useState('')
  const [busy, setBusy] = useState(false)

  const [checking, setChecking] = useState(false)
  useEffect(() => {
    api.worker.status().then((x) => {
      setW((x as Worker) ?? null)
      // Opening Settings checks a paired worker: online or not, right now.
      if (x) {
        setChecking(true)
        api.worker.check().then((y) => setW((y as Worker) ?? null)).catch(() => {}).finally(() => setChecking(false))
      }
    }).catch(() => {})
    api.worker.clientFingerprint().then((f) => setMe(f || '')).catch(() => {})
  }, [])
  // PowerShell actions may become available (or not) through the worker.
  const refreshCatalog = () => setCache('catalog.rev', Date.now())

  const run = async (fn: () => Promise<unknown>, ok?: string): Promise<boolean> => {
    setBusy(true)
    try {
      const x = await fn()
      setW((x as Worker) ?? null)
      refreshCatalog()
      if (ok) toast('ok', ok)
      return true
    } catch (e) { toast('err', errMessage(e)); return false } finally { setBusy(false) }
  }

  return (
    <Card title={t('worker.title')}>
      <p className="mb-3 text-xs text-[var(--text-faint)]">{t('worker.intro')}</p>
      {w ? (
        <div className="flex flex-col gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2">
          <div className="flex items-center gap-2">
            <Server size={15} className="text-[var(--text-faint)]" />
            <div className="min-w-0 flex-1">
              <div className="truncate text-sm font-medium">{w.name || w.addr}</div>
              <div className="truncate text-xs text-[var(--text-faint)]">{w.addr}{w.version ? ` · ${w.version}` : ''}</div>
            </div>
            <span className="flex items-center gap-1.5 text-xs" data-worker-state>
              <span className={`h-2 w-2 rounded-full ${checking ? 'bg-[var(--text-faint)]' : w.online ? 'bg-[var(--ok)]' : 'bg-[var(--danger)]'}`} />
              {checking ? t('worker.checking') : w.online ? t('worker.online') : t('worker.offline')}
            </span>
          </div>
          <div className="flex flex-wrap gap-1">
            {(w.families ?? []).map((f) => <Badge key={f} kind="ok">{familyName(f)}</Badge>)}
            {(w.unavailable ?? []).map((f) => <Badge key={f} kind="warn">{familyName(f)} · {t('worker.notReady')}</Badge>)}
          </div>
          <p className="text-xs text-[var(--text-faint)]">
            {when(w.lastSeen) && <>{t('worker.lastSeen', { when: when(w.lastSeen) })} · </>}
            {t('worker.pairedAt', { when: when(w.pairedAt) })}
          </p>
          {!checking && !w.online && w.lastError && <p className="text-xs text-[var(--danger)]">{w.lastError}</p>}
          <p className="break-all font-mono text-[10px] text-[var(--text-faint)]">{t('worker.pinned')} {w.fingerprint}</p>
          {(w.families ?? []).length === 0 && <p className="text-xs text-[var(--warn)]">{t('worker.noFamilies')}</p>}
          <div className="flex gap-2">
            <Button variant="subtle" disabled={busy || checking} onClick={() => run(() => api.worker.check())}><RefreshCw size={14} /> {t('worker.check')}</Button>
            <Button variant="ghost" disabled={busy} onClick={() => run(async () => { await api.worker.unpair(); return null })}><Unlink size={14} /> {t('worker.unpair')}</Button>
          </div>
        </div>
      ) : (
        <div className="flex flex-col gap-2">
          <ol className="list-decimal pl-5 text-xs text-[var(--text-dim)]">
            <li>{t('worker.step1')} <code className="font-mono">SwissKnifeGraph worker serve --pair</code></li>
            <li>{t('worker.step2')}</li>
          </ol>
          <Field label={t('worker.address')}><Input value={addr} onChange={(e) => setAddr(e.target.value)} placeholder="srv-ps01:8743" /></Field>
          <Field label={t('worker.code')}><Input value={code} onChange={(e) => setCode(e.target.value)} placeholder="XXXX-XXXX-XXXX-XXXX" className="font-mono" /></Field>
          <Button variant="primary" disabled={busy || !addr.trim() || !code.trim()}
            onClick={() => run(() => api.worker.pair(addr, code), t('worker.paired')).then((ok) => { if (ok) setCode('') })}>
            <Link2 size={14} /> {t('worker.pair')}
          </Button>
        </div>
      )}
      {me && <p className="mt-3 break-all font-mono text-[10px] text-[var(--text-faint)]">{t('worker.thisApp')} {me}</p>}
    </Card>
  )
}
