import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Inbox, Send, CalendarDays, CalendarPlus } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Textarea } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
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

  const userField = (
    <Field label={t('common.user')}>
      <EntityPicker value={user} onChange={setUser} load={loadUsers} placeholder={t('mail.pickUser')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'read', label: t('mail.tileRead'), hint: t('mail.hintRead'), icon: <Inbox size={16} />, variant: 'primary',
      note: <p>{t('mail.noteRead')}</p>,
      panel: (
        <TaskForm>
          {userField}
          <Field label={t('mail.folder')}><Input value={folder} onChange={(e) => setFolder(e.target.value)} placeholder="inbox" /></Field>
          <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.mail.list(user, folder, 25))}>
            <Inbox size={15} /> {t('common.run')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'send', label: t('mail.tileSend'), hint: t('mail.hintSend'), icon: <Send size={16} />, variant: 'danger', write: true,
      note: <p className="text-[var(--warn)]">{t('mail.noteSend')}</p>,
      panel: (
        <TaskForm>
          {userField}
          <Field label={t('mail.subject')}><Input value={mail.subject} onChange={(e) => setMail({ ...mail, subject: e.target.value })} /></Field>
          <Field label={t('mail.to')}><Input value={mail.to} onChange={(e) => setMail({ ...mail, to: e.target.value })} placeholder="a@contoso.com, b@contoso.com" /></Field>
          <Field label={t('mail.body')}><Textarea rows={4} value={mail.body} onChange={(e) => setMail({ ...mail, body: e.target.value })} /></Field>
          <Button variant="danger" disabled={readOnly || !user || !mail.to}
            onClick={() => askConfirm(user, (c) => doWrite('send', () => api.mail.send(user, mail.subject, mail.body, mail.to.split(',').map((s) => s.trim()).filter(Boolean), c), t('mail.sendAs')))}>
            <Send size={15} /> {t('mail.sendAs')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'calendar', label: t('mail.tileCalendar'), hint: t('mail.hintCalendar'), icon: <CalendarDays size={16} />,
      panel: (
        <TaskForm>
          {userField}
          <Button variant="primary" disabled={!user} onClick={() => res.run(() => api.calendar.list(user, 25))}>
            <CalendarDays size={15} /> {t('mail.listEvents')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'createEvent', label: t('mail.tileCreateEvent'), hint: t('mail.hintCreateEvent'), icon: <CalendarPlus size={16} />, write: true,
      note: <p>{t('mail.noteCreateEvent')}</p>,
      panel: (
        <TaskForm>
          {userField}
          <Field label={t('mail.subject')}><Input value={ev.subject} onChange={(e) => setEv({ ...ev, subject: e.target.value })} /></Field>
          <div className="grid grid-cols-2 gap-2">
            <Field label={t('mail.start')}><Input value={ev.start} onChange={(e) => setEv({ ...ev, start: e.target.value })} placeholder="2026-07-28T10:00:00" /></Field>
            <Field label={t('mail.end')}><Input value={ev.end} onChange={(e) => setEv({ ...ev, end: e.target.value })} placeholder="2026-07-28T11:00:00" /></Field>
          </div>
          <Field label={t('mail.timezone')}><Input value={ev.tz} onChange={(e) => setEv({ ...ev, tz: e.target.value })} /></Field>
          <Field label={t('mail.attendees')}><Input value={ev.attendees} onChange={(e) => setEv({ ...ev, attendees: e.target.value })} placeholder="a@contoso.com" /></Field>
          <Field label={t('mail.body')}><Textarea rows={2} value={ev.body} onChange={(e) => setEv({ ...ev, body: e.target.value })} /></Field>
          <Button variant="primary" disabled={readOnly || !user || !ev.subject}
            onClick={() => doWrite('createEvent', async () => res.setData(await api.calendar.createEvent(user, ev.subject, ev.body, ev.start, ev.end, ev.tz, ev.attendees.split(',').map((s) => s.trim()).filter(Boolean))), t('mail.createEvent'))}>
            <CalendarPlus size={15} /> {t('mail.createEvent')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="mail"
        title={t('nav.mail')}
        subtitle={t('mail.subtitle')}
        actions={actions}
        status={status}
        busy={res.loading}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </>
  )
}
