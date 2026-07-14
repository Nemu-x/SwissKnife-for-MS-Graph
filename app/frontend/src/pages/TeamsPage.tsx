import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { ListTree, Hash, UserPlus, UserMinus, Wand2, Plus } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { EntityPicker } from '../components/EntityPicker'
import { loadTeams, loadChannels, loadGroups } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function TeamsPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [user, setUser] = useState('')
  const [teamId, setTeamId] = useState('')
  const [channelId, setChannelId] = useState('')
  const [upn, setUpn] = useState('')
  const [owner, setOwner] = useState(false)
  const [ch, setCh] = useState({ name: '', desc: '', type: 'standard', owner: '' })
  const [groupId, setGroupId] = useState('')

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    {
      id: 'browse', label: 'Browse', icon: <ListTree size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Field label={t('common.user')}><UpnInput value={user} onChange={setUser} /></Field>
          <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.teams.joined(user))}><ListTree size={15} /> Joined teams</Button>
          <Field label="Team"><EntityPicker value={teamId} onChange={setTeamId} load={loadTeams} placeholder="Pick a team…" /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={!teamId} onClick={() => res.run(() => api.teams.channels(teamId))}><Hash size={15} /> Channels</Button>
            <Button variant="subtle" disabled={!teamId} onClick={() => res.run(() => api.teams.teamMembers(teamId))}>Members</Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'membership', label: 'Membership', icon: <UserPlus size={15} />,
      panel: (
        <DrawerForm>
          <Field label="Team"><EntityPicker value={teamId} onChange={setTeamId} load={loadTeams} placeholder="Pick a team…" /></Field>
          <Field label="Channel (optional)">
            <EntityPicker value={channelId} onChange={setChannelId} load={loadChannels(teamId)} reloadKey={teamId} placeholder="Pick a channel…" />
          </Field>
          <Field label={t('common.user')}><UpnInput value={upn} onChange={setUpn} /></Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={owner} onChange={(e) => setOwner(e.target.checked)} /> {t('common.owner')}
          </label>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !teamId || !upn}
              onClick={() => doWrite(() => channelId ? api.teams.addChannelMember(teamId, channelId, upn, owner) : api.teams.addTeamMember(teamId, upn, owner), t('common.add'))}>
              <UserPlus size={15} /> {t('common.add')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !teamId || !upn}
              onClick={() => doWrite(() => channelId ? api.teams.removeChannelMember(teamId, channelId, upn) : api.teams.removeTeamMember(teamId, upn), t('common.remove'))}>
              <UserMinus size={15} /> {t('common.remove')}
            </Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'create', label: 'Create / Teamify', icon: <Plus size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Field label="Team"><EntityPicker value={teamId} onChange={setTeamId} load={loadTeams} placeholder="Pick a team…" /></Field>
          <Input placeholder="Channel name" value={ch.name} onChange={(e) => setCh({ ...ch, name: e.target.value })} />
          <Input placeholder="Description" value={ch.desc} onChange={(e) => setCh({ ...ch, desc: e.target.value })} />
          <Select value={ch.type} onChange={(e) => setCh({ ...ch, type: e.target.value })} className="w-full">
            <option value="standard">standard</option><option value="private">private</option><option value="shared">shared</option>
          </Select>
          {ch.type !== 'standard' && <Input placeholder="Owner UPN (required)" value={ch.owner} onChange={(e) => setCh({ ...ch, owner: e.target.value })} />}
          <Button variant="primary" disabled={readOnly || !teamId || !ch.name}
            onClick={() => doWrite(async () => res.setData(await api.teams.createChannel(teamId, ch.name, ch.desc, ch.type, ch.owner)), t('common.create'))}>
            <Plus size={15} /> Channel
          </Button>
          <div className="my-1 border-t border-[var(--border)]" />
          <Field label="Group → Team"><EntityPicker value={groupId} onChange={setGroupId} load={loadGroups} placeholder="Pick a group…" /></Field>
          <Button variant="subtle" disabled={readOnly || !groupId} onClick={() => doWrite(async () => res.setData(await api.teams.teamify(groupId)), 'Teamify')}>
            <Wand2 size={15} /> Teamify
          </Button>
        </DrawerForm>
      ),
    },
  ]

  return <ActionPage title={t('nav.teams')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} />} />
}
