import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { MessageSquare, Users2, UserPlus, UserMinus, Plus } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input } from '../components/ui'
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

  return (
    <Page title={t('nav.chats')}>
      <TwoPane
        controls={
          <>
            <Card title={t('nav.chats')}>
              <div className="flex flex-col gap-2">
                <Field label={t('common.user')}><Input value={user} onChange={(e) => setUser(e.target.value)} /></Field>
                <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.chats.list(user, 0))}>
                  <MessageSquare size={15} /> List chats
                </Button>
                <Field label="Chat ID"><Input value={chatId} onChange={(e) => setChatId(e.target.value)} /></Field>
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" disabled={!chatId} onClick={() => res.run(() => api.chats.messages(chatId, 50))}>Messages</Button>
                  <Button variant="subtle" disabled={!chatId} onClick={() => res.run(() => api.chats.members(chatId))}>
                    <Users2 size={15} /> Members
                  </Button>
                </div>
              </div>
            </Card>

            <Card title="Membership">
              <div className="flex flex-col gap-2">
                <Field label={t('common.user')}><Input value={upn} onChange={(e) => setUpn(e.target.value)} /></Field>
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" disabled={readOnly || !chatId || !upn}
                    onClick={() => doWrite(() => api.chats.addMember(chatId, upn, false), t('common.add'))}>
                    <UserPlus size={15} /> {t('common.add')}
                  </Button>
                  <Button variant="subtle" disabled={readOnly || !chatId || !upn}
                    onClick={() => doWrite(() => api.chats.removeMember(chatId, upn), t('common.remove'))}>
                    <UserMinus size={15} /> {t('common.remove')}
                  </Button>
                </div>
              </div>
            </Card>

            <Card title="Create group chat">
              <div className="flex flex-col gap-2">
                <Input placeholder="Topic" value={topic} onChange={(e) => setTopic(e.target.value)} />
                <Input placeholder="Members (comma-separated UPNs)" value={members} onChange={(e) => setMembers(e.target.value)} />
                <Button variant="primary" disabled={readOnly || !members}
                  onClick={() => doWrite(async () => res.setData(await api.chats.createGroup(topic, members.split(',').map((s) => s.trim()).filter(Boolean))), t('common.create'))}>
                  <Plus size={15} /> {t('common.create')}
                </Button>
              </div>
            </Card>
          </>
        }
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </Page>
  )
}
