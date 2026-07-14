import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ShieldCheck, Users2, UserPlus, UserMinus } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function RolesPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [roleId, setRoleId] = useState('')
  const [upn, setUpn] = useState('')

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    { id: 'list', label: t('roles.listRoles'), icon: <ShieldCheck size={15} />, variant: 'primary', onClick: () => res.run(() => api.roles.list()) },
    {
      id: 'members', label: t('roles.members'), icon: <Users2 size={15} />,
      panel: (
        <DrawerForm>
          <Field label={t('roles.roleId')}><Input value={roleId} onChange={(e) => setRoleId(e.target.value)} /></Field>
          <Button variant="subtle" disabled={!roleId} onClick={() => res.run(() => api.roles.members(roleId))}><Users2 size={15} /> {t('roles.members')}</Button>
          <Field label={t('common.user')}><UpnInput value={upn} onChange={setUpn} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !roleId || !upn} onClick={() => doWrite(() => api.roles.addMember(roleId, upn), t('roles.addMember'))}><UserPlus size={15} /> {t('roles.addMember')}</Button>
            <Button variant="danger" disabled={readOnly || !roleId || !upn}
              onClick={() => askConfirm(upn, (c) => doWrite(() => api.roles.removeMember(roleId, upn, c), t('roles.removeMember')))}>
              <UserMinus size={15} /> {t('roles.removeMember')}
            </Button>
          </div>
        </DrawerForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <ActionPage title={t('roles.title')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} />} />
    </>
  )
}
