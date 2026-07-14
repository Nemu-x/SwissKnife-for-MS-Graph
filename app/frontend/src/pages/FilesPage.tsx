import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { FolderTree, Search, Upload, Download, Trash2, Link2, Building2 } from 'lucide-react'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input, Select } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function FilesPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()

  const [ownerType, setOwnerType] = useState<'user' | 'site'>('user')
  const [ownerId, setOwnerId] = useState('')
  const [siteSearch, setSiteSearch] = useState('')
  const [itemId, setItemId] = useState('')
  const [itemName, setItemName] = useState('')
  const [folder, setFolder] = useState('')
  const [query, setQuery] = useState('')
  const [progress, setProgress] = useState<{ name: string; pct: number } | null>(null)

  useEffect(() => {
    const off = EventsOn('transfer:progress', (d: any) => {
      const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
      setProgress({ name: d.name, pct })
      if (pct >= 100) setTimeout(() => setProgress(null), 1500)
    })
    return () => off()
  }, [])

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  return (
    <Page title={t('nav.files')}>
      {confirmElement}
      <TwoPane
        controls={
          <>
            <Card title="Drive">
              <div className="flex flex-col gap-2">
                <Select value={ownerType} onChange={(e) => setOwnerType(e.target.value as any)} className="w-full">
                  <option value="user">OneDrive (user)</option>
                  <option value="site">SharePoint (site)</option>
                </Select>
                {ownerType === 'site' && (
                  <div className="flex gap-2">
                    <Input value={siteSearch} onChange={(e) => setSiteSearch(e.target.value)} placeholder="Site search" />
                    <Button variant="subtle" onClick={() => res.run(() => api.drive.sites(siteSearch))}><Building2 size={15} /></Button>
                  </div>
                )}
                <Field label={ownerType === 'user' ? t('common.user') : 'Site ID'}>
                  <Input value={ownerId} onChange={(e) => setOwnerId(e.target.value)} />
                </Field>
                <Button variant="primary" disabled={!ownerId} onClick={() => res.run(() => api.drive.listRoot(ownerType, ownerId))}>
                  <FolderTree size={15} /> List root
                </Button>
                <div className="flex gap-2">
                  <Input value={query} onChange={(e) => setQuery(e.target.value)} placeholder={t('common.search')} />
                  <Button variant="subtle" disabled={!ownerId} onClick={() => res.run(() => api.drive.search(ownerType, ownerId, query))}>
                    <Search size={15} />
                  </Button>
                </div>
              </div>
            </Card>

            <Card title="Item ops">
              <div className="flex flex-col gap-2">
                <Field label="Item ID"><Input value={itemId} onChange={(e) => setItemId(e.target.value)} /></Field>
                <Input placeholder="Suggested name (download)" value={itemName} onChange={(e) => setItemName(e.target.value)} />
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" disabled={!ownerId || !itemId}
                    onClick={() => doWrite(async () => { const p = await api.drive.download(ownerType, ownerId, itemId, itemName || 'download'); if (p) toast('ok', p) }, t('common.run'))}>
                    <Download size={15} /> Download
                  </Button>
                  <Button variant="subtle" disabled={readOnly || !ownerId || !itemId}
                    onClick={() => doWrite(() => api.drive.createLink(ownerType, ownerId, itemId, 'view', 'organization'), 'link')}>
                    <Link2 size={15} /> Link
                  </Button>
                </div>
                <Input placeholder="Upload to folder ('' = root)" value={folder} onChange={(e) => setFolder(e.target.value)} />
                <Button variant="primary" disabled={readOnly || !ownerId}
                  onClick={() => doWrite(async () => res.setData(await api.drive.upload(ownerType, ownerId, folder)), 'upload')}>
                  <Upload size={15} /> Upload file…
                </Button>
                <Button variant="danger" disabled={readOnly || !ownerId || !itemId}
                  onClick={() => askConfirm(itemId, (c) => doWrite(() => api.drive.delete(ownerType, ownerId, itemId, c), t('common.delete')))}>
                  <Trash2 size={15} /> {t('common.delete')}
                </Button>
              </div>
            </Card>

            {progress && (
              <div className="rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-3">
                <div className="mb-1 flex justify-between text-xs text-[var(--text-dim)]">
                  <span className="truncate">{progress.name}</span><span>{progress.pct}%</span>
                </div>
                <div className="h-1.5 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                  <div className="h-full bg-[var(--accent)] transition-all" style={{ width: `${progress.pct}%` }} />
                </div>
              </div>
            )}
          </>
        }
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </Page>
  )
}
