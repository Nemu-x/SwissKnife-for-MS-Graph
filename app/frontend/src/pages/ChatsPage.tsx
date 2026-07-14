import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { MessageSquare, Users2, UserPlus, UserMinus, Plus } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function ChatsPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [user, setUser] = useState('')
  const [chatId, setChatId] = useState('')
  const [upn, setUpn] = useState('')
  const [topic, setTopic] = useState('')
  const [members, setMembers] = useState('')

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    {
      id: 'browse', label: 'Browse', icon: <MessageSquare size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Field label={t('common.user')}><EntityPicker value={user} onChange={setUser} load={loadUsers} placeholder="Pick a user…" /></Field>
          <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.chats.list(user, 0))}><MessageSquare size={15} /> List chats</Button>
          <Field label="Chat ID"><Input value={chatId} onChange={(e) => setChatId(e.target.value)} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={!chatId} onClick={() => res.run(() => api.chats.messages(chatId, 50))}>Messages</Button>
            <Button variant="subtle" disabled={!chatId} onClick={() => res.run(() => api.chats.members(chatId))}><Users2 size={15} /> Members</Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'membership', label: 'Membership', icon: <UserPlus size={15} />,
      panel: (
        <DrawerForm>
          <Field label="Chat ID"><Input value={chatId} onChange={(e) => setChatId(e.target.value)} /></Field>
          <Field label={t('common.user')}><UpnInput value={upn} onChange={setUpn} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !chatId || !upn} onClick={() => doWrite(() => api.chats.addMember(chatId, upn, false), t('common.add'))}><UserPlus size={15} /> {t('common.add')}</Button>
            <Button variant="subtle" disabled={readOnly || !chatId || !upn} onClick={() => doWrite(() => api.chats.removeMember(chatId, upn), t('common.remove'))}><UserMinus size={15} /> {t('common.remove')}</Button>
          </div>
        </DrawerForm>
      ),
    },
    {
      id: 'create', label: t('common.create'), icon: <Plus size={15} />, variant: 'primary',
      panel: (
        <DrawerForm>
          <Input placeholder="Topic" value={topic} onChange={(e) => setTopic(e.target.value)} />
          <Input placeholder="Members (comma-separated UPNs)" value={members} onChange={(e) => setMembers(e.target.value)} />
          <Button variant="primary" disabled={readOnly || !members}
            onClick={() => doWrite(async () => res.setData(await api.chats.createGroup(topic, members.split(',').map((s) => s.trim()).filter(Boolean))), t('common.create'))}>
            <Plus size={15} /> {t('common.create')}
          </Button>
        </DrawerForm>
      ),
    },
  ]

  return <ActionPage title={t('nav.chats')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} />} />
}
