import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { List, UserCheck, Plus, Minus } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers, loadSkus } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function LicensingPage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [target, setTarget] = useState('')
  const [skuId, setSkuId] = useState('')
  const { status, busy: writing, doWrite } = useTaskStatus()

  const userField = (
    <Field label={t('common.user')}>
      <EntityPicker value={target} onChange={setTarget} load={loadUsers} placeholder={t('licensing.pickUser')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'skus', label: t('licensing.tileSkus'), hint: t('licensing.hintSkus'), icon: <List size={16} />, variant: 'primary',
      onClick: () => res.run(() => api.licensing.skus()),
    },
    {
      id: 'userLicenses', label: t('licensing.tileUserLicenses'), hint: t('licensing.hintUserLicenses'), icon: <UserCheck size={16} />,
      panel: (
        <TaskForm>
          {userField}
          <Button variant="primary" disabled={!target} onClick={() => res.run(() => api.users.licenseDetails(target))}>
            <UserCheck size={15} /> {t('licensing.userLicenses')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'assign', label: t('licensing.tileAssign'), hint: t('licensing.hintAssign'), icon: <Plus size={16} />, variant: 'primary', write: true,
      note: (
        <>
          <p>{t('licensing.noteAssign')}</p>
          <p className="text-[var(--warn)]">{t('licensing.noteRemove')}</p>
        </>
      ),
      panel: (
        <TaskForm>
          {userField}
          <Field label={t('licensing.sku')} hint={t('licensing.skuHint')}>
            <EntityPicker value={skuId} onChange={setSkuId} load={loadSkus} placeholder={t('licensing.pickSku')} />
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !target || !skuId}
              onClick={() => doWrite('assign', () => api.licensing.assign(target, [skuId], []), t('licensing.assigned'))}>
              <Plus size={15} /> {t('common.add')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !target || !skuId}
              onClick={() => doWrite('assign', () => api.licensing.assign(target, [], [skuId]), t('licensing.removed'))}>
              <Minus size={15} /> {t('common.remove')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
  ]

  return (
    <TaskPage
      pageId="licensing"
      title={t('nav.licensing')}
      subtitle={t('licensing.subtitle')}
      actions={actions}
      status={status}
      busy={res.loading || writing}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
