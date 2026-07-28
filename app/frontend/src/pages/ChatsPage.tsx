import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { MessageSquare, Users2, UserPlus, UserMinus, Plus } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers, loadChats } from '../lib/pickers'
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
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  const doWrite = async (id: string, fn: () => Promise<any>, ok: string) => {
    try {
      await fn()
      setStatus((s) => ({ ...s, [id]: { ok: true, text: ok, at: Date.now() } }))
      toast('ok', ok)
    } catch (e) {
      const m = errMessage(e)
      setStatus((s) => ({ ...s, [id]: { ok: false, text: m, at: Date.now() } }))
      toast('err', m)
    }
  }

  // Chats are listed per owner: every tile needs to know whose chats to look at.
  const ownerField = (
    <Field label={t('chats.owner')}>
      <EntityPicker value={user} onChange={setUser} load={loadUsers} placeholder={t('chats.pickUser')} />
    </Field>
  )
  const chatField = (
    <Field label={t('chats.chat')}>
      <EntityPicker value={chatId} onChange={setChatId} load={loadChats(user)} reloadKey={user}
        placeholder={user ? t('chats.pickChat') : t('chats.pickUserFirst')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'read', label: t('chats.tileRead'), hint: t('chats.hintRead'), icon: <MessageSquare size={16} />, variant: 'primary',
      note: <p>{t('chats.noteRead')}</p>,
      panel: (
        <TaskForm>
          {ownerField}
          <Button variant="subtle" disabled={!user} onClick={() => res.run(() => api.chats.list(user, 0))}>
            <MessageSquare size={15} /> {t('chats.listChats')}
          </Button>
          {chatField}
          <Button variant="primary" disabled={!chatId} onClick={() => res.run(() => api.chats.messages(chatId, 50))}>
            {t('chats.messages')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'members', label: t('chats.tileMembers'), hint: t('chats.hintMembers'), icon: <Users2 size={16} />, write: true,
      panel: (
        <TaskForm>
          {ownerField}
          {chatField}
          <Button variant="subtle" disabled={!chatId} onClick={() => res.run(() => api.chats.members(chatId))}>
            <Users2 size={15} /> {t('chats.members')}
          </Button>
          <Field label={t('chats.memberToAdd')}>
            <EntityPicker value={upn} onChange={setUpn} load={loadUsers} placeholder={t('chats.pickUser')} />
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !chatId || !upn} onClick={() => doWrite('members', () => api.chats.addMember(chatId, upn, false), t('common.add'))}>
              <UserPlus size={15} /> {t('common.add')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !chatId || !upn} onClick={() => doWrite('members', () => api.chats.removeMember(chatId, upn), t('common.remove'))}>
              <UserMinus size={15} /> {t('common.remove')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'create', label: t('chats.tileCreate'), hint: t('chats.hintCreate'), icon: <Plus size={16} />, write: true,
      note: <p>{t('chats.noteCreate')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('chats.topic')}><Input value={topic} onChange={(e) => setTopic(e.target.value)} /></Field>
          <Field label={t('chats.membersCsv')}>
            <Input value={members} onChange={(e) => setMembers(e.target.value)} placeholder="a@contoso.com, b@contoso.com" />
          </Field>
          <Button variant="primary" disabled={readOnly || !members}
            onClick={() => doWrite('create', async () => res.setData(await api.chats.createGroup(topic, members.split(',').map((s) => s.trim()).filter(Boolean))), t('common.create'))}>
            <Plus size={15} /> {t('common.create')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <TaskPage
      pageId="chats"
      title={t('nav.chats')}
      subtitle={t('chats.subtitle')}
      actions={actions}
      status={status}
      busy={res.loading}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
