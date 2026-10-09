import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Download, RefreshCw, ShieldCheck, Workflow, FileCode2 } from 'lucide-react'
import { Button, Badge, Input, Spinner } from './ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { localized } from './CatalogAction'

type HubPack = {
  name: string; version: string; category: string; kind: 'workflow' | 'script' | 'mixed'
  title: Record<string, string>; description?: Record<string, string>; author?: string
  signed: boolean; installed?: string; update: boolean
}

const CATEGORIES = ['all', 'security', 'exchange', 'hr', 'teams', 'intune', 'reports', 'other']

// The Action Hub: community packs to browse and install. An installed pack
// still needs a trusted signature or the operator's review before it runs.
export function HubCatalog({ onInstalled }: { onInstalled: (list: unknown) => void }) {
  const { t } = useTranslation()
  const { toast } = useStore()
  const [view, setView] = useState<{ url: string; packs: HubPack[] } | null>(null)
  const [error, setError] = useState('')
  const [cat, setCat] = useState('all')
  const [busy, setBusy] = useState('')
  const [hub, setHub] = useState('')

  const load = () => {
    setError('')
    setView(null)
    api.packs.hubCatalog()
      .then((v: any) => { setView({ url: v?.url ?? '', packs: v?.packs ?? [] }); setHub(v?.url ?? '') })
      .catch((e) => setError(errMessage(e)))
  }
  useEffect(load, [])

  const install = async (name: string) => {
    setBusy(name)
    try {
      onInstalled(await api.packs.hubInstall(name))
      toast('ok', t('hub.installed'))
      load()
    } catch (e) { toast('err', errMessage(e)) } finally { setBusy('') }
  }
  const saveHub = async () => {
    try { await api.packs.setHub(hub); load() } catch (e) { toast('err', errMessage(e)) }
  }

  const shown = (view?.packs ?? []).filter((p) => cat === 'all' || p.category === cat)
  return (
    <div className="flex flex-col gap-3">
      <p className="text-xs text-[var(--text-faint)]">{t('hub.intro')}</p>
      <div className="flex flex-wrap gap-1">
        {CATEGORIES.map((c) => (
          <button key={c} onClick={() => setCat(c)}
            className={`rounded-full px-2.5 py-0.5 text-xs ${cat === c ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'border border-[var(--border)] text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]'}`}>
            {t(`hub.cat.${c}`)}
          </button>
        ))}
      </div>
      {!view && !error && <Spinner />}
      {error && (
        <div className="flex flex-col gap-2">
          <p className="text-sm text-[var(--warn)]">{t('hub.unavailable')}</p>
          <p className="text-xs text-[var(--text-faint)]">{error}</p>
          <Button variant="subtle" onClick={load}><RefreshCw size={14} /> {t('hub.retry')}</Button>
        </div>
      )}
      {view && shown.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('hub.empty')}</p>}
      {shown.map((p) => (
        <div key={p.name} className="flex gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2.5">
          <span className="mt-0.5 text-[var(--accent)]">{p.kind === 'workflow' ? <Workflow size={18} /> : <FileCode2 size={18} />}</span>
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-sm font-medium">{localized(p.title) || p.name}</span>
              <Badge kind="neutral">{t(`hub.kind.${p.kind}`)}</Badge>
              {p.signed && <Badge kind="ok"><ShieldCheck size={11} /> {t('hub.signed')}</Badge>}
              <span className="text-xs text-[var(--text-faint)]">{p.version}{p.author ? ` · ${p.author}` : ''}</span>
            </div>
            {p.description && <p className="mt-0.5 text-xs text-[var(--text-dim)]">{localized(p.description)}</p>}
          </div>
          <div className="shrink-0 self-center">
            {p.installed && !p.update
              ? <Badge kind="ok">{t('hub.isInstalled')}</Badge>
              : (
                <Button variant={p.update ? 'subtle' : 'primary'} disabled={!!busy} onClick={() => install(p.name)}>
                  {busy === p.name ? <Spinner /> : <Download size={14} />} {p.update ? t('hub.update', { v: p.version }) : t('hub.install')}
                </Button>
              )}
          </div>
        </div>
      ))}
      <details className="text-xs text-[var(--text-faint)]">
        <summary className="cursor-pointer">{t('hub.address')}</summary>
        <div className="mt-2 flex gap-2">
          <Input value={hub} onChange={(e) => setHub(e.target.value)} className="min-w-0 flex-1 font-mono text-xs" />
          <Button variant="subtle" onClick={saveHub}>{t('common.save')}</Button>
        </div>
      </details>
    </div>
  )
}
