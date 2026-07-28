import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ListTree, Hash, UserPlus, UserMinus, Wand2, Plus, Users2, MapPin } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { EntityPicker } from '../components/EntityPicker'
import { loadTeams, loadChannels, loadMembershipChannels, loadGroups, loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function TeamsPage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  // One "who" and one "where" for the whole page: the person and team picked in
  // one tile stay picked in the next.
  const [user, setUser] = useState('')
  const [teamId, setTeamId] = useState('')
  // Two channel pickers with different option sets: browsing lists every
  // channel, membership only the ones that have their own members. Sharing one
  // state let a standard channel leak into the add-member form, where Graph
  // then rejects the call.
  const [channelId, setChannelId] = useState('')
  const [memberChannelId, setMemberChannelId] = useState('')
  const [owner, setOwner] = useState(false)
  const [ch, setCh] = useState({ name: '', desc: '', type: 'standard', owner: '' })
  const [groupId, setGroupId] = useState('')
  const { status, busy: writing, doWrite } = useTaskStatus()

  const userField = (
    <Field label={t('common.user')}>
      <EntityPicker value={user} onChange={setUser} load={loadUsers} placeholder={t('teams.pickUser')} />
    </Field>
  )
  const teamField = (
    <Field label={t('teams.team')}>
      <EntityPicker value={teamId} onChange={setTeamId} load={loadTeams} placeholder={t('teams.pickTeam')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'addToTeam', label: t('teams.tileAddToTeam'), hint: t('teams.hintAddToTeam'),
      icon: <UserPlus size={16} />, variant: 'primary', write: true,
      note: <p>{t('teams.noteAddToTeam')}</p>,
      panel: (
        <TaskForm>
          {userField}
          {teamField}
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={owner} onChange={(e) => setOwner(e.target.checked)} /> {t('common.owner')}
          </label>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !teamId || !user}
              onClick={() => doWrite('addToTeam', () => api.teams.addTeamMember(teamId, user, owner), t('teams.addToTeam'))}>
              <UserPlus size={15} /> {t('teams.addToTeam')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !teamId || !user}
              onClick={() => doWrite('addToTeam', () => api.teams.removeTeamMember(teamId, user), t('teams.removeFromTeam'))}>
              <UserMinus size={15} /> {t('teams.removeFromTeam')}
            </Button>
          </div>
          <Button variant="ghost" disabled={!user} onClick={() => res.run(() => api.teams.joined(user))}>
            <ListTree size={15} /> {t('teams.checkMembership')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'addToChannel', label: t('teams.tileAddToChannel'), hint: t('teams.hintAddToChannel'),
      icon: <Hash size={16} />, variant: 'primary', write: true,
      note: (
        <>
          <p className="text-[var(--warn)]">{t('teams.orderHint')}</p>
          <p>{t('teams.standardHint')}</p>
        </>
      ),
      panel: (
        <TaskForm>
          {userField}
          {teamField}
          <Field label={t('teams.channel')}>
            <EntityPicker value={memberChannelId} onChange={setMemberChannelId} load={loadMembershipChannels(teamId)} reloadKey={teamId}
              placeholder={teamId ? t('teams.pickChannel') : t('teams.pickTeamFirst')} />
          </Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={owner} onChange={(e) => setOwner(e.target.checked)} /> {t('common.owner')}
          </label>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !teamId || !memberChannelId || !user}
              onClick={() => doWrite('addToChannel', () => api.teams.addChannelMember(teamId, memberChannelId, user, owner), t('teams.addToChannel'))}>
              <UserPlus size={15} /> {t('teams.addToChannel')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !teamId || !memberChannelId || !user}
              onClick={() => doWrite('addToChannel', () => api.teams.removeChannelMember(teamId, memberChannelId, user), t('teams.removeFromChannel'))}>
              <UserMinus size={15} /> {t('teams.removeFromChannel')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'whereIs', label: t('teams.tileWhereIs'), hint: t('teams.hintWhereIs'),
      icon: <MapPin size={16} />,
      panel: (
        <TaskForm>
          {userField}
          <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.teams.joined(user))}>
            <ListTree size={15} /> {t('teams.joinedTeams')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'teamContents', label: t('teams.tileContents'), hint: t('teams.hintContents'),
      icon: <Users2 size={16} />,
      panel: (
        <TaskForm>
          {teamField}
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={!teamId} onClick={() => res.run(() => api.teams.channels(teamId))}><Hash size={15} /> {t('teams.channels')}</Button>
            <Button variant="subtle" disabled={!teamId} onClick={() => res.run(() => api.teams.teamMembers(teamId))}>{t('teams.members')}</Button>
          </div>
          <Field label={t('teams.channel')}>
            <EntityPicker value={channelId} onChange={setChannelId} load={loadChannels(teamId)} reloadKey={teamId}
              placeholder={teamId ? t('teams.pickChannel') : t('teams.pickTeamFirst')} />
          </Field>
          <Button variant="subtle" disabled={!teamId || !channelId} onClick={() => res.run(() => api.teams.channelMembers(teamId, channelId))}>
            <Users2 size={15} /> {t('teams.channelMembers')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'createChannel', label: t('teams.tileCreateChannel'), hint: t('teams.hintCreateChannel'),
      icon: <Plus size={16} />, write: true,
      note: <p>{t('teams.noteCreateChannel')}</p>,
      panel: (
        <TaskForm>
          {teamField}
          <Field label={t('teams.channelName')}>
            <Input value={ch.name} onChange={(e) => setCh({ ...ch, name: e.target.value })} />
          </Field>
          <Field label={t('teams.channelDesc')}>
            <Input value={ch.desc} onChange={(e) => setCh({ ...ch, desc: e.target.value })} />
          </Field>
          <Field label={t('teams.channelType')}>
            <Select value={ch.type} onChange={(e) => setCh({ ...ch, type: e.target.value })} className="w-full">
              <option value="standard">standard</option><option value="private">private</option><option value="shared">shared</option>
            </Select>
          </Field>
          {ch.type !== 'standard' && (
            <Field label={t('teams.channelOwner')}>
              <UpnInput value={ch.owner} onChange={(v) => setCh({ ...ch, owner: v })} />
            </Field>
          )}
          <Button variant="primary" disabled={readOnly || !teamId || !ch.name || (ch.type !== 'standard' && !ch.owner)}
            onClick={() => doWrite('createChannel', async () => res.setData(await api.teams.createChannel(teamId, ch.name, ch.desc, ch.type, ch.owner)), t('teams.createChannel'))}>
            <Plus size={15} /> {t('teams.createChannel')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'teamify', label: t('teams.tileTeamify'), hint: t('teams.hintTeamify'),
      icon: <Wand2 size={16} />, write: true,
      note: <p>{t('teams.noteTeamify')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('teams.groupToTeam')}>
            <EntityPicker value={groupId} onChange={setGroupId} load={loadGroups} placeholder={t('teams.pickGroup')} />
          </Field>
          <Button variant="primary" disabled={readOnly || !groupId}
            onClick={() => doWrite('teamify', async () => res.setData(await api.teams.teamify(groupId)), t('teams.teamify'))}>
            <Wand2 size={15} /> {t('teams.teamify')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <TaskPage
      pageId="teams"
      title={t('nav.teams')}
      subtitle={t('teams.subtitle')}
      actions={actions}
      status={status}
      busy={res.loading || writing}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
