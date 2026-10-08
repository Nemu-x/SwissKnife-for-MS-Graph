import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { FolderOpen, RefreshCw, ShieldCheck, ShieldOff, KeyRound, Trash2 } from 'lucide-react'
import { Card, Button, Badge, Input } from './ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { localized } from './CatalogAction'

type Pack = {
  name: string; version: string; author: string; description: string; dir: string
  status: 'signed' | 'trusted' | 'untrusted' | 'changed' | 'invalid'
  signer?: string; error?: string; digest: string
  actions: { id: string; label: Record<string, string>; page: string; danger: string; module: string }[]
}

const KIND: Record<Pack['status'], 'ok' | 'warn' | 'danger' | 'neutral'> = {
  signed: 'ok', trusted: 'ok', untrusted: 'neutral', changed: 'warn', invalid: 'danger',
}

// Community action packs (ADR-008): what is installed, whether it is trusted,
// and the signing keys trusted besides the project's own.
export function PacksCard() {
  const { t } = useTranslation()
  const { toast, setCache } = useStore()
  const [list, setList] = useState<Pack[] | null>(null)
  const [keys, setKeys] = useState<string[]>([])
  const [newKey, setNewKey] = useState('')
  // The catalog changes with the trust: pages fetch it again.
  const changed = (l: unknown) => { setList((l ?? []) as Pack[]); setCache('catalog.rev', Date.now()) }

  const load = () => {
    api.packs.list().then((l) => setList((l ?? []) as Pack[])).catch((e) => toast('err', errMessage(e)))
    api.packs.keys().then((k) => setKeys(k ?? [])).catch(() => {})
  }
  useEffect(load, [])

  const act = async (fn: () => Promise<unknown>) => {
    try { changed(await fn()) } catch (e) { toast('err', errMessage(e)) }
  }
  const addKey = async () => {
    try { setKeys((await api.packs.addKey(newKey)) ?? []); setNewKey(''); load(); setCache('catalog.rev', Date.now()) } catch (e) { toast('err', errMessage(e)) }
  }
  const removeKey = async (k: string) => {
    try { setKeys((await api.packs.removeKey(k)) ?? []); load(); setCache('catalog.rev', Date.now()) } catch (e) { toast('err', errMessage(e)) }
  }

  return (
    <Card title={t('packs.title')}>
      <p className="mb-3 text-xs text-[var(--text-faint)]">{t('packs.intro')}</p>
      <div className="mb-3 flex gap-2">
        <Button variant="subtle" onClick={() => api.packs.openFolder().catch((e) => toast('err', errMessage(e)))}><FolderOpen size={14} /> {t('packs.openFolder')}</Button>
        <Button variant="subtle" onClick={load}><RefreshCw size={14} /> {t('packs.reload')}</Button>
      </div>
      {list?.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('packs.none')}</p>}
      <div className="flex flex-col gap-2">
        {list?.map((p) => (
          <div key={p.dir} className="rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2">
            <div className="flex items-center gap-2">
              <div className="min-w-0 flex-1">
                <div className="truncate text-sm font-medium">{p.name} <span className="text-xs text-[var(--text-faint)]">{p.version}</span></div>
                <div className="truncate text-xs text-[var(--text-faint)]">{p.author || p.dir}</div>
              </div>
              <Badge kind={KIND[p.status]}>{t(`packs.status.${p.status}`)}</Badge>
            </div>
            {p.description && <p className="mt-1 text-xs text-[var(--text-dim)]">{p.description}</p>}
            {p.error && <p className="mt-1 text-xs text-[var(--danger)]">{p.error}</p>}
            {p.signer && <p className="mt-1 text-xs text-[var(--text-faint)]">{t('packs.signedBy', { key: p.signer })}</p>}
            {p.actions.length > 0 && (
              <ul className="mt-1 list-disc pl-5 text-xs text-[var(--text-dim)]">
                {p.actions.map((a) => (
                  <li key={a.id}>{localized(a.label) || a.id} · {t(`nav.${a.page}`, { defaultValue: a.page })} · {t(`packs.danger.${a.danger}`)} · {a.module === 'teams' ? 'Teams PS' : 'Exchange PS'}</li>
                ))}
              </ul>
            )}
            {(p.status === 'untrusted' || p.status === 'changed') && (
              <div className="mt-2 flex flex-col gap-1.5">
                <p className="text-xs text-[var(--warn)]">{t('packs.reviewWarn', { dir: p.dir })}</p>
                <p className="break-all font-mono text-[10px] text-[var(--text-faint)]">SHA-256 {p.digest}</p>
                <Button variant="primary" onClick={() => act(() => api.packs.trust(p.name, p.digest))}><ShieldCheck size={14} /> {t('packs.trust')}</Button>
              </div>
            )}
            {p.status === 'trusted' && (
              <Button variant="ghost" className="mt-2" onClick={() => act(() => api.packs.untrust(p.name))}><ShieldOff size={14} /> {t('packs.untrust')}</Button>
            )}
          </div>
        ))}
      </div>

      <p className="mb-1 mt-4 text-xs font-medium text-[var(--text-dim)]">{t('packs.keys')}</p>
      <p className="mb-2 text-xs text-[var(--text-faint)]">{t('packs.keysHint')}</p>
      {keys.map((k) => (
        <div key={k} className="mb-1 flex items-center gap-2">
          <KeyRound size={13} className="shrink-0 text-[var(--text-faint)]" />
          <span className="min-w-0 flex-1 truncate font-mono text-xs">{k}</span>
          <Button variant="ghost" className="!px-2 !py-1" onClick={() => removeKey(k)}><Trash2 size={13} /></Button>
        </div>
      ))}
      <div className="flex gap-2">
        <Input value={newKey} onChange={(e) => setNewKey(e.target.value)} placeholder="RWS…" className="min-w-0 flex-1 font-mono text-xs" />
        <Button variant="subtle" disabled={!newKey.trim()} onClick={addKey}>{t('common.add')}</Button>
      </div>
    </Card>
  )
}
