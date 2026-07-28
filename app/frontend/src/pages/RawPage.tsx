import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Play, Star, History, Trash2 } from 'lucide-react'
import { Page } from '../components/Layout'
import { ResultView } from '../components/ResultView'
import { Button, Card, Select, Input, Textarea, Spinner } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

interface RawItem { method: string; path: string; body: string }

const EXAMPLES: RawItem[] = [
  { method: 'GET', path: '/me', body: '' },
  { method: 'GET', path: '/organization', body: '' },
  { method: 'GET', path: "/users?$top=5&$select=displayName,userPrincipalName", body: '' },
  { method: 'GET', path: '/subscribedSkus', body: '' },
]

function load(key: string): RawItem[] {
  try { return JSON.parse(localStorage.getItem(key) || '[]') } catch { return [] }
}
function store(key: string, v: RawItem[]) { localStorage.setItem(key, JSON.stringify(v.slice(0, 30))) }

export function RawPage() {
  const { t } = useTranslation()
  const { toast } = useStore()
  const res = useAsync<GraphObject>()
  const [item, setItem] = useState<RawItem>({ method: 'GET', path: '/me', body: '' })
  const [history, setHistory] = useState<RawItem[]>(load('raw.history'))
  const [favorites, setFavorites] = useState<RawItem[]>(load('raw.favorites'))

  const send = async () => {
    await res.run(() => api.raw.send(item.method, item.path, item.body))
    const h = [item, ...history.filter((x) => !(x.method === item.method && x.path === item.path))]
    setHistory(h); store('raw.history', h)
  }
  const addFav = () => {
    const f = [item, ...favorites]; setFavorites(f); store('raw.favorites', f); toast('ok', t('raw.addFavorite'))
  }
  const clearHist = () => { setHistory([]); store('raw.history', []) }
  const removeFav = (i: number) => { const f = favorites.filter((_, idx) => idx !== i); setFavorites(f); store('raw.favorites', f) }

  const bodyDisabled = item.method === 'GET' || item.method === 'DELETE'

  return (
    <Page title={t('nav.raw')}>
      <div className="grid h-full grid-cols-1 gap-4 lg:grid-cols-[minmax(320px,400px)_1fr]">
        <div className="flex flex-col gap-4 overflow-auto">
          <Card title={t('nav.raw')}>
            <div className="flex flex-col gap-2">
              <div className="flex gap-2">
                <Select value={item.method} onChange={(e) => setItem({ ...item, method: e.target.value })}>
                  {['GET', 'POST', 'PATCH', 'PUT', 'DELETE'].map((m) => <option key={m}>{m}</option>)}
                </Select>
                <Input value={item.path} onChange={(e) => setItem({ ...item, path: e.target.value })} placeholder="/me" />
              </div>
              <Textarea rows={6} disabled={bodyDisabled} value={item.body}
                onChange={(e) => setItem({ ...item, body: e.target.value })}
                placeholder={bodyDisabled ? '—' : t('raw.body')} />
              <div className="flex gap-2">
                <Button variant="primary" className="flex-1" disabled={res.loading} onClick={send}>
                  {res.loading ? <Spinner /> : <Play size={15} />} {res.loading ? t('common.working') : t('common.run')}
                </Button>
                <Button variant="ghost" onClick={addFav}><Star size={15} /></Button>
              </div>
            </div>
          </Card>

          <Card title={t('raw.examples')}>
            <div className="flex flex-col gap-1">
              {EXAMPLES.map((e, i) => (
                <button key={i} onClick={() => setItem(e)}
                  className="rounded-md px-2 py-1 text-left text-xs text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]">
                  <span className="font-mono text-[var(--accent)]">{e.method}</span> {e.path}
                </button>
              ))}
            </div>
          </Card>

          {favorites.length > 0 && (
            <Card title={t('raw.favorites')}>
              <div className="flex flex-col gap-1">
                {favorites.map((e, i) => (
                  <div key={i} className="flex items-center gap-1">
                    <button onClick={() => setItem(e)} className="flex-1 truncate rounded-md px-2 py-1 text-left text-xs hover:bg-[var(--bg-elev-2)]">
                      <span className="font-mono text-[var(--accent)]">{e.method}</span> {e.path}
                    </button>
                    <button onClick={() => removeFav(i)} className="text-[var(--text-faint)] hover:text-[var(--danger)]"><Trash2 size={13} /></button>
                  </div>
                ))}
              </div>
            </Card>
          )}

          {history.length > 0 && (
            <Card title={t('raw.history')} actions={<button onClick={clearHist} className="text-[var(--text-faint)] hover:text-[var(--danger)]"><Trash2 size={14} /></button>}>
              <div className="flex flex-col gap-1">
                {history.map((e, i) => (
                  <button key={i} onClick={() => setItem(e)} className="flex items-center gap-1 truncate rounded-md px-2 py-1 text-left text-xs hover:bg-[var(--bg-elev-2)]">
                    <History size={12} className="shrink-0 text-[var(--text-faint)]" />
                    <span className="font-mono text-[var(--accent)]">{e.method}</span> {e.path}
                  </button>
                ))}
              </div>
            </Card>
          )}
        </div>

        <div className="min-h-[300px] overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--bg-elev)]">
          <ResultView data={res.data} loading={res.loading} error={res.error} />
        </div>
      </div>
    </Page>
  )
}
