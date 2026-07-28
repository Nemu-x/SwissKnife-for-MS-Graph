import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { FolderTree, Search, Upload, Download, Trash2, Link2 } from 'lucide-react'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
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
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  useEffect(() => {
    const off = EventsOn('transfer:progress', (d: any) => {
      const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
      setProgress({ name: d.name, pct })
      if (pct >= 100) setTimeout(() => setProgress(null), 1500)
    })
    return () => off()
  }, [])

  const doWrite = async (id: string, fn: () => Promise<any>, ok: string) => {
    try {
      await fn()
      setStatus((s) => ({ ...s, [id]: { ok: true, text: ok, at: Date.now() } }))
      toast('ok', ok)
    } catch (e) {
      const m = errMessage(e)
      setStatus((s) => ({ ...s, [id]: { ok: false, text: m, at: Date.now() } }))
      toast('err', m)
    }
  }

  // Every tile works against one drive: a user's OneDrive or a SharePoint site.
  const driveField = (
    <>
      <Select value={ownerType} onChange={(e) => { setOwnerType(e.target.value as any); setOwnerId('') }} className="w-full">
        <option value="user">{t('files.oneDrive')}</option>
        <option value="site">{t('files.sharePoint')}</option>
      </Select>
      <Field label={ownerType === 'user' ? t('common.user') : t('files.site')}>
        <EntityPicker value={ownerId} onChange={setOwnerId}
          load={ownerType === 'user' ? loadUsers : loadSites} reloadKey={ownerType}
          placeholder={ownerType === 'user' ? t('files.pickUser') : t('files.pickSite')} />
      </Field>
    </>
  )

  const actions: TaskAction[] = [
    {
      id: 'drive', label: t('files.tileBrowse'), hint: t('files.hintBrowse'), icon: <FolderTree size={16} />, variant: 'primary',
      panel: (
        <TaskForm>
          {driveField}
          <Button variant="primary" disabled={!ownerId} onClick={() => res.run(() => api.drive.listRoot(ownerType, ownerId))}>
            <FolderTree size={15} /> {t('files.listRoot')}
          </Button>
          <Field label={t('common.search')}>
            <div className="flex gap-2">
              <Input value={query} onChange={(e) => setQuery(e.target.value)} />
              <Button variant="subtle" disabled={!ownerId} onClick={() => res.run(() => api.drive.search(ownerType, ownerId, query))}><Search size={15} /></Button>
            </div>
          </Field>
        </TaskForm>
      ),
    },
    {
      id: 'upload', label: t('files.tileUpload'), hint: t('files.hintUpload'), icon: <Upload size={16} />, write: true,
      note: <p>{t('files.noteUpload')}</p>,
      panel: (
        <TaskForm>
          {driveField}
          <Field label={t('files.uploadFolder')}><Input value={folder} onChange={(e) => setFolder(e.target.value)} /></Field>
          <Button variant="primary" disabled={readOnly || !ownerId}
            onClick={() => doWrite('upload', async () => res.setData(await api.drive.upload(ownerType, ownerId, folder)), t('files.upload'))}>
            <Upload size={15} /> {t('files.upload')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'item', label: t('files.tileItem'), hint: t('files.hintItem'), icon: <Download size={16} />, write: true,
      note: <p>{t('files.itemIdHint')}</p>,
      panel: (
        <TaskForm>
          {driveField}
          <Field label={t('files.itemId')}><Input value={itemId} onChange={(e) => setItemId(e.target.value)} /></Field>
          <Field label={t('files.suggestedName')}><Input value={itemName} onChange={(e) => setItemName(e.target.value)} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={!ownerId || !itemId}
              onClick={() => doWrite('item', async () => { const p = await api.drive.download(ownerType, ownerId, itemId, itemName || 'download'); if (p) toast('ok', p) }, t('files.download'))}>
              <Download size={15} /> {t('files.download')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !ownerId || !itemId}
              onClick={() => doWrite('item', () => api.drive.createLink(ownerType, ownerId, itemId, 'view', 'organization'), t('files.link'))}>
              <Link2 size={15} /> {t('files.link')}
            </Button>
          </div>
          <Button variant="danger" disabled={readOnly || !ownerId || !itemId}
            onClick={() => askConfirm(itemId, (c) => doWrite('item', () => api.drive.delete(ownerType, ownerId, itemId, c), t('common.delete')))}>
            <Trash2 size={15} /> {t('common.delete')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="files"
        title={t('nav.files')}
        subtitle={t('files.subtitle')}
        actions={actions}
        status={status}
        busy={res.loading || !!progress}
        busyLabel={progress ? `${progress.name} — ${progress.pct}%` : undefined}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error || !!progress}
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
            {/* "Use this ID" fills the item field, so a listed file can be acted on. */}
            <div className="min-h-0 flex-1"><ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setItemId} /></div>
          </div>
        }
      />
    </>
  )
}
