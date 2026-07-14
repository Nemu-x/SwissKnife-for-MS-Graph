import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Inbox, Send, CalendarDays, CalendarPlus } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input, Textarea } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function MailPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [user, setUser] = useState('')
  const [folder, setFolder] = useState('inbox')
  const [mail, setMail] = useState({ subject: '', body: '', to: '' })
  const [ev, setEv] = useState({ subject: '', body: '', start: '', end: '', tz: 'UTC', attendees: '' })

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  return (
    <Page title={t('nav.mail')}>
      {confirmElement}
      <TwoPane
        controls={
          <>
            <Card title="Mail">
              <div className="flex flex-col gap-2">
                <Field label={t('common.user')}><Input value={user} onChange={(e) => setUser(e.target.value)} /></Field>
                <div className="flex gap-2">
                  <Input value={folder} onChange={(e) => setFolder(e.target.value)} placeholder="inbox" />
                  <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.mail.list(user, folder, 25))}>
                    <Inbox size={15} />
                  </Button>
                </div>
                <Input placeholder="Subject" value={mail.subject} onChange={(e) => setMail({ ...mail, subject: e.target.value })} />
                <Input placeholder="To (comma-separated)" value={mail.to} onChange={(e) => setMail({ ...mail, to: e.target.value })} />
                <Textarea rows={3} placeholder="Body" value={mail.body} onChange={(e) => setMail({ ...mail, body: e.target.value })} />
                <Button variant="danger" disabled={readOnly || !user || !mail.to}
                  onClick={() => askConfirm(user, (c) => doWrite(() => api.mail.send(user, mail.subject, mail.body, mail.to.split(',').map((s) => s.trim()).filter(Boolean), c), t('common.run')))}>
                  <Send size={15} /> Send as user
                </Button>
              </div>
            </Card>

            <Card title="Calendar">
              <div className="flex flex-col gap-2">
                <Button variant="subtle" disabled={!user} onClick={() => res.run(() => api.calendar.list(user, 25))}>
                  <CalendarDays size={15} /> List events
                </Button>
                <Input placeholder="Subject" value={ev.subject} onChange={(e) => setEv({ ...ev, subject: e.target.value })} />
                <div className="grid grid-cols-2 gap-2">
                  <Input placeholder="Start ISO" value={ev.start} onChange={(e) => setEv({ ...ev, start: e.target.value })} />
                  <Input placeholder="End ISO" value={ev.end} onChange={(e) => setEv({ ...ev, end: e.target.value })} />
                </div>
                <Input placeholder="Time zone" value={ev.tz} onChange={(e) => setEv({ ...ev, tz: e.target.value })} />
                <Input placeholder="Attendees (comma-separated)" value={ev.attendees} onChange={(e) => setEv({ ...ev, attendees: e.target.value })} />
                <Textarea rows={2} placeholder="Body" value={ev.body} onChange={(e) => setEv({ ...ev, body: e.target.value })} />
                <Button variant="primary" disabled={readOnly || !user || !ev.subject}
                  onClick={() => doWrite(async () => res.setData(await api.calendar.createEvent(user, ev.subject, ev.body, ev.start, ev.end, ev.tz, ev.attendees.split(',').map((s) => s.trim()).filter(Boolean))), t('common.create'))}>
                  <CalendarPlus size={15} /> Create event
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
