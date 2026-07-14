import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { FolderTree, Search, Upload, Download, Trash2, Link2 } from 'lucide-react'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers, loadSites } from '../lib/pickers'
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

  const actions: Action[] = [
    {
      id: 'drive', label: 'Drive', icon: <FolderTree size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Select value={ownerType} onChange={(e) => { setOwnerType(e.target.value as any); setOwnerId('') }} className="w-full">
            <option value="user">OneDrive (user)</option>
            <option value="site">SharePoint (site)</option>
          </Select>
          <Field label={ownerType === 'user' ? t('common.user') : 'SharePoint site'}>
            <EntityPicker value={ownerId} onChange={setOwnerId}
              load={ownerType === 'user' ? loadUsers : loadSites} reloadKey={ownerType}
              placeholder={ownerType === 'user' ? 'Pick a user…' : 'Pick a site…'} />
          </Field>
          <Button variant="primary" disabled={!ownerId} onClick={() => res.run(() => api.drive.listRoot(ownerType, ownerId))}><FolderTree size={15} /> List root</Button>
          <div className="flex gap-2">
            <Input value={query} onChange={(e) => setQuery(e.target.value)} placeholder={t('common.search')} />
            <Button variant="subtle" disabled={!ownerId} onClick={() => res.run(() => api.drive.search(ownerType, ownerId, query))}><Search size={15} /></Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'item', label: 'Item ops', icon: <Upload size={15} />,
      panel: (
        <DrawerForm>
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
          <Button variant="primary" disabled={readOnly || !ownerId} onClick={() => doWrite(async () => res.setData(await api.drive.upload(ownerType, ownerId, folder)), 'upload')}>
            <Upload size={15} /> Upload file…
          </Button>
          <Button variant="danger" disabled={readOnly || !ownerId || !itemId}
            onClick={() => askConfirm(itemId, (c) => doWrite(() => api.drive.delete(ownerType, ownerId, itemId, c), t('common.delete')))}>
            <Trash2 size={15} /> {t('common.delete')}
          </Button>
        </DrawerForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <ActionPage
        title={t('nav.files')}
        actions={actions}
        result={
          <div className="flex h-full flex-col">
            {progress && (
              <div className="border-b border-[var(--border)] p-3">
                <div className="mb-1 flex justify-between text-xs text-[var(--text-dim)]">
                  <span className="truncate">{progress.name}</span><span>{progress.pct}%</span>
                </div>
                <div className="h-1.5 overflow-hidden rounded bg-[var(--bg-elev-2)]">
                  <div className="h-full bg-[var(--accent)] transition-all" style={{ width: `${progress.pct}%` }} />
                </div>
              </div>
            )}
            <div className="min-h-0 flex-1"><ResultView data={res.data} loading={res.loading} error={res.error} /></div>
          </div>
        }
      />
    </>
  )
}
