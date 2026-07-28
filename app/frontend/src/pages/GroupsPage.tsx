import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Users2, UserPlus, Plus } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadGroups, loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function GroupsPage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [groupId, setGroupId] = useState('')
  const [upn, setUpn] = useState('')
  const [create, setCreate] = useState({ name: '', desc: '', nick: '', owner: '' })
  const { status, busy: writing, doWrite } = useTaskStatus()

  const listGroups = () => res.run(() => api.groups.list(search, 0))

  const groupField = (
    <Field label={t('groups.group')}>
      <EntityPicker value={groupId} onChange={setGroupId} load={loadGroups} placeholder={t('groups.pickGroup')} />
    </Field>
  )

  const actions: TaskAction[] = [
    { id: 'list', label: t('groups.tileList'), hint: t('groups.hintList'), icon: <Search size={16} />, onClick: listGroups },
    {
      id: 'members', label: t('groups.tileMembers'), hint: t('groups.hintMembers'), icon: <Users2 size={16} />, variant: 'primary',
      panel: (
        <TaskForm>
          {groupField}
          <Button variant="primary" disabled={!groupId} onClick={() => res.run(() => api.groups.members(groupId))}>
            <Users2 size={15} /> {t('roles.members')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'add', label: t('groups.tileAdd'), hint: t('groups.hintAdd'), icon: <UserPlus size={16} />, variant: 'primary', write: true,
      note: <p>{t('groups.noteAdd')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('common.user')}>
            <EntityPicker value={upn} onChange={setUpn} load={loadUsers} placeholder={t('groups.pickUser')} />
          </Field>
          {groupField}
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !groupId || !upn} onClick={() => doWrite('add', () => api.groups.addMember(groupId, upn), t('groups.addMember'))}>
              <UserPlus size={15} /> {t('groups.addMember')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !groupId || !upn} onClick={() => doWrite('add', () => api.groups.addOwner(groupId, upn), t('groups.addOwner'))}>
              <UserPlus size={15} /> {t('groups.addOwner')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'create', label: t('groups.tileCreate'), hint: t('groups.hintCreate'), icon: <Plus size={16} />, write: true,
      note: <p>{t('groups.noteCreate')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('users.displayName')}><Input value={create.name} onChange={(e) => setCreate({ ...create, name: e.target.value })} /></Field>
          <Field label={t('groups.description')}><Input value={create.desc} onChange={(e) => setCreate({ ...create, desc: e.target.value })} /></Field>
          <Field label={t('users.mailNickname')}><Input value={create.nick} onChange={(e) => setCreate({ ...create, nick: e.target.value })} /></Field>
          <Field label={t('groups.ownerOptional')}><Input value={create.owner} onChange={(e) => setCreate({ ...create, owner: e.target.value })} placeholder="owner@contoso.com" /></Field>
          <Button variant="primary" disabled={readOnly || !create.name || !create.nick}
            onClick={() => doWrite('create', async () => res.setData(await api.groups.createM365(create.name, create.desc, create.nick, create.owner)), t('common.create'))}>
            <Plus size={15} /> {t('common.create')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <TaskPage
      pageId="groups"
      title={t('nav.groups')}
      subtitle={t('groups.subtitle')}
      search={{ value: search, onChange: setSearch, onSubmit: listGroups, placeholder: t('common.search') }}
      actions={actions}
      status={status}
      busy={res.loading || writing}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setGroupId} />}
    />
  )
}
