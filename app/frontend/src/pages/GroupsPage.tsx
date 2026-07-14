import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Users2, UserPlus, Plus } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function GroupsPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [groupId, setGroupId] = useState('')
  const [upn, setUpn] = useState('')
  const [create, setCreate] = useState({ name: '', desc: '', nick: '', owner: '' })

  const listGroups = () => res.run(() => api.groups.list(search, 0))
  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    { id: 'search', label: t('common.search'), icon: <Search size={15} />, variant: 'primary', onClick: listGroups },
    {
      id: 'members', label: t('roles.members'), icon: <Users2 size={15} />,
      panel: (
        <DrawerForm>
          <Field label="Group ID"><Input value={groupId} onChange={(e) => setGroupId(e.target.value)} /></Field>
          <Button variant="subtle" disabled={!groupId} onClick={() => res.run(() => api.groups.members(groupId))}><Users2 size={15} /> {t('roles.members')}</Button>
          <Field label={t('common.user')}><UpnInput value={upn} onChange={setUpn} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !groupId || !upn} onClick={() => doWrite(() => api.groups.addOwner(groupId, upn), t('common.owner'))}><UserPlus size={15} /> {t('common.owner')}</Button>
            <Button variant="subtle" disabled={readOnly || !groupId || !upn} onClick={() => doWrite(() => api.groups.addMember(groupId, upn), t('common.member'))}><UserPlus size={15} /> {t('common.member')}</Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'create', label: t('common.create'), icon: <Plus size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Input placeholder="Display name" value={create.name} onChange={(e) => setCreate({ ...create, name: e.target.value })} />
          <Input placeholder="Description" value={create.desc} onChange={(e) => setCreate({ ...create, desc: e.target.value })} />
          <Input placeholder="Mail nickname" value={create.nick} onChange={(e) => setCreate({ ...create, nick: e.target.value })} />
          <Input placeholder="Owner UPN (optional)" value={create.owner} onChange={(e) => setCreate({ ...create, owner: e.target.value })} />
          <Button variant="primary" disabled={readOnly || !create.name || !create.nick}
            onClick={() => doWrite(async () => res.setData(await api.groups.createM365(create.name, create.desc, create.nick, create.owner)), t('common.create'))}>
            <Plus size={15} /> {t('common.create')}
          </Button>
        </DrawerForm>
      ),
    },
  ]

  return (
    <ActionPage
      title={t('nav.groups')}
      search={{ value: search, onChange: setSearch, onSubmit: listGroups, placeholder: t('common.search') }}
      actions={actions}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
