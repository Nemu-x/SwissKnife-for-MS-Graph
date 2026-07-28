import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ShieldCheck, Users2, UserPlus } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadRoles, loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function RolesPage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [roleId, setRoleId] = useState('')
  const [upn, setUpn] = useState('')
  const { status, busy: writing, doWrite } = useTaskStatus()

  const roleField = (
    <Field label={t('roles.role')}>
      <EntityPicker value={roleId} onChange={setRoleId} load={loadRoles} placeholder={t('roles.pickRole')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'list', label: t('roles.tileList'), hint: t('roles.hintList'), icon: <ShieldCheck size={16} />, variant: 'primary',
      onClick: () => res.run(() => api.roles.list()),
    },
    {
      id: 'members', label: t('roles.tileMembers'), hint: t('roles.hintMembers'), icon: <Users2 size={16} />,
      panel: (
        <TaskForm>
          {roleField}
          <Button variant="primary" disabled={!roleId} onClick={() => res.run(() => api.roles.members(roleId))}>
            <Users2 size={15} /> {t('roles.members')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'grant', label: t('roles.tileGrant'), hint: t('roles.hintGrant'), icon: <UserPlus size={16} />, variant: 'primary', write: true,
      note: <p>{t('roles.noteGrant')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('common.user')}>
            <EntityPicker value={upn} onChange={setUpn} load={loadUsers} placeholder={t('roles.pickUser')} />
          </Field>
          {roleField}
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !roleId || !upn} onClick={() => doWrite('grant', () => api.roles.addMember(roleId, upn), t('roles.addMember'))}>
              <UserPlus size={15} /> {t('roles.addMember')}
            </Button>
            <Button variant="danger" disabled={readOnly || !roleId || !upn}
              onClick={() => askConfirm(upn, (c) => doWrite('grant', () => api.roles.removeMember(roleId, upn, c), t('roles.removeMember')))}>
              {t('roles.removeMember')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="roles"
        title={t('roles.title')}
        subtitle={t('roles.subtitle')}
        actions={actions}
        status={status}
        busy={res.loading || writing}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setRoleId} />}
      />
    </>
  )
}
