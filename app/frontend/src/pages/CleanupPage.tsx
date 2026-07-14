import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Trash2, Building2 } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Field, Input, Select, Badge, Spinner, ErrorNote } from '../components/ui'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

export function CleanupPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()

  const [ownerType, setOwnerType] = useState<'user' | 'site'>('user')
  const [ownerId, setOwnerId] = useState('')
  const [groups, setGroups] = useState<services.DupGroup[] | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const scan = async () => {
    setBusy(true); setError(null); setGroups(null)
    try { setGroups(await api.cleanup.findDuplicates(ownerType, ownerId)) } catch (e) { setError(errMessage(e)) } finally { setBusy(false) }
  }

  // Delete every copy except the first in each group.
  const deleteExtras = () => {
    if (!groups) return
    const ids = groups.flatMap((g) => g.items.slice(1).map((i) => i.id))
    if (ids.length === 0) return
    askConfirm('DELETE', async (c) => {
      try {
        const r = await api.cleanup.deleteItems(ownerType, ownerId, ids, c)
        toast('ok', `${r.deleted} deleted`)
        scan()
      } catch (e) { toast('err', errMessage(e)) }
    }, t('cleanup.deleteExtras'))
  }

  const totalWasted = (groups || []).reduce((a, g) => a + g.wasted, 0)

  return (
    <Page title={t('cleanup.title')} subtitle={t('cleanup.subtitle')}>
      {confirmElement}
      <Card title={t('cleanup.title')} className="mb-4">
        <div className="flex flex-wrap items-end gap-3">
          <label className="flex flex-col gap-1">
            <span className="text-xs text-[var(--text-dim)]">Drive</span>
            <Select value={ownerType} onChange={(e) => setOwnerType(e.target.value as any)} className="w-48">
              <option value="user">OneDrive (user)</option>
              <option value="site">SharePoint (site)</option>
            </Select>
          </label>
          <Field label={ownerType === 'user' ? t('common.user') : 'Site ID'}>
            <Input value={ownerId} onChange={(e) => setOwnerId(e.target.value)} />
          </Field>
          <Button variant="primary" disabled={!ownerId || busy} onClick={scan}>
            {busy ? <Spinner /> : ownerType === 'site' ? <Building2 size={15} /> : <Search size={15} />}
            {busy ? t('cleanup.scanning') : t('cleanup.scan')}
          </Button>
          {groups && groups.length > 0 && (
            <Button variant="danger" disabled={readOnly} onClick={deleteExtras} className="ml-auto">
              <Trash2 size={15} /> {t('cleanup.deleteExtras')}
            </Button>
          )}
        </div>
      </Card>

      {error && <ErrorNote>{error}</ErrorNote>}

      {groups && groups.length === 0 && (
        <p className="text-sm text-[var(--ok)]">{t('cleanup.noDupes')}</p>
      )}

      {groups && groups.length > 0 && (
        <>
          <div className="mb-3 text-sm text-[var(--text-dim)]">
            {t('cleanup.totalWasted', { size: humanBytes(totalWasted) })}
          </div>
          <div className="flex flex-col gap-2">
            {groups.map((g, i) => (
              <div key={i} className="rounded-xl border border-[var(--border)] bg-[var(--bg-elev)] p-3">
                <div className="flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="truncate text-sm font-medium">{g.name}</div>
                    <div className="text-xs text-[var(--text-faint)]">{humanBytes(g.size)} each · {t('cleanup.copies', { n: g.count })}</div>
                  </div>
                  <Badge kind="warn">{t('cleanup.wasted')}: {humanBytes(g.wasted)}</Badge>
                </div>
                <div className="mt-2 flex flex-col gap-1">
                  {g.items.map((it, j) => (
                    <div key={it.id} className="flex items-center gap-2 text-xs">
                      <span className={j === 0 ? 'text-[var(--ok)]' : 'text-[var(--text-faint)]'}>{j === 0 ? '✓ keep' : '✗ extra'}</span>
                      <span className="truncate text-[var(--text-dim)]">{it.path}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </Page>
  )
}
