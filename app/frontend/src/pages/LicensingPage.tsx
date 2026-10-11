import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { List, UserCheck } from 'lucide-react'
import { catalogTile } from '../components/CatalogAction'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { api, type GraphObject } from '../lib/api'

export function LicensingPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [target, setTarget] = useState('')

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
    catalogTile('license.assign'),
  ]

  return (
    <TaskPage
      pageId="licensing"
      title={t('nav.licensing')}
      subtitle={t('licensing.subtitle')}
      actions={actions}
      busy={res.loading}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
