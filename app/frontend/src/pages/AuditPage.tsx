import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { LogIn, FileClock, ShieldX, MailSearch, ListTree, PlugZap } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select, ErrorNote } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { UpnInput } from '../components/UpnInput'
import { loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

const PERIODS = [1, 7, 30]
// Exchange answers at most 10 days per message-trace query (90 days back in total).
const TRACE_PERIODS = [1, 2, 5, 10]
// Typed into the confirm modal before the one-time tenant write is sent.
const ENABLE_TRACE_TARGET = 'message-trace'

export function AuditPage() {
  const { t } = useTranslation()
  const { toast, readOnly } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[]>()
  const [upn, setUpn] = useState('')
  const [days, setDays] = useState(7)
  const [top, setTop] = useState(50)
  const [actor, setActor] = useState('')

  // Message trace: "the email did not arrive, where did it go?"
  const [sender, setSender] = useState('')
  const [recipient, setRecipient] = useState('')
  const [traceDays, setTraceDays] = useState(2)
  const [traceTop, setTraceTop] = useState(100)
  const [traceId, setTraceId] = useState('')
  const [traceRecipient, setTraceRecipient] = useState('')
  // null = not checked; false = the Microsoft service principal is missing.
  const [tracePrereq, setTracePrereq] = useState<boolean | null>(null)
  const [enabling, setEnabling] = useState(false)

  const runTrace = async () => {
    const rows = await res.run(() => api.mailTrace.trace(sender, recipient, traceDays, traceTop))
    if (rows !== undefined) {
      setTracePrereq(true)
      return
    }
    // The call failed: find out whether it is the tenant prerequisite rather than
    // a permission or filter problem. The check itself may be forbidden — then
    // stay silent and leave the raw error on screen.
    try {
      setTracePrereq(await api.mailTrace.prerequisite())
    } catch {
      /* no Application.Read.All — cannot tell */
    }
  }

  const enableTrace = () =>
    askConfirm(
      ENABLE_TRACE_TARGET,
      async () => {
        setEnabling(true)
        try {
          await api.mailTrace.provision()
          toast('ok', t('mailTrace.enabled'))
          setTracePrereq(true)
          await runTrace()
        } catch (e) {
          toast('err', errMessage(e))
        } finally {
          setEnabling(false)
        }
      },
      t('mailTrace.enableTitle'),
    )

  // "Use this ID" on a trace row fills the details form with that row's id and recipient.
  const useTraceRow = (id: string) => {
    setTraceId(id)
    const row = (res.data ?? []).find((r) => String(r.id) === id)
    if (row?.recipientAddress) setTraceRecipient(String(row.recipientAddress))
  }

  const period = (
    <Field label={t('audit.period')}>
      <Select value={days} onChange={(e) => setDays(Number(e.target.value))} className="w-full">
        {PERIODS.map((d) => <option key={d} value={d}>{t('audit.lastDays', { n: d })}</option>)}
        <option value={0}>{t('audit.anyTime')}</option>
      </Select>
    </Field>
  )
  const limit = (
    <Field label={t('audit.limit')}>
      <Input type="number" value={top} onChange={(e) => setTop(Math.max(1, Number(e.target.value) || 50))} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'whyNoSignIn', label: t('audit.tileWhyFailed'), hint: t('audit.hintWhyFailed'),
      icon: <ShieldX size={16} />, variant: 'primary',
      note: <p>{t('audit.noteWhyFailed')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('common.user')}>
            <EntityPicker value={upn} onChange={setUpn} load={loadUsers} placeholder={t('users.pickUser')} />
          </Field>
          {period}
          {limit}
          <Button variant="primary" disabled={!upn} onClick={() => res.run(() => api.auditQuery.signIns(upn, days, true, top))}>
            <ShieldX size={15} /> {t('audit.failedOnly')}
          </Button>
          <Button variant="subtle" disabled={!upn} onClick={() => res.run(() => api.auditQuery.signIns(upn, days, false, top))}>
            <LogIn size={15} /> {t('audit.allSignIns')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'trace', label: t('mailTrace.tile'), hint: t('mailTrace.hint'),
      icon: <MailSearch size={16} />, variant: 'primary',
      note: <p>{t('mailTrace.note')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('mailTrace.sender')} hint={t('mailTrace.addressHint')}>
            <UpnInput value={sender} onChange={setSender} placeholder="sender@example.com" />
          </Field>
          <Field label={t('mailTrace.recipient')}>
            <UpnInput value={recipient} onChange={setRecipient} placeholder="recipient@contoso.com" />
          </Field>
          <Field label={t('audit.period')}>
            <Select value={traceDays} onChange={(e) => setTraceDays(Number(e.target.value))} className="w-full">
              {TRACE_PERIODS.map((d) => <option key={d} value={d}>{t('audit.lastDays', { n: d })}</option>)}
            </Select>
          </Field>
          <Field label={t('audit.limit')}>
            <Input type="number" value={traceTop} onChange={(e) => setTraceTop(Math.max(1, Number(e.target.value) || 100))} />
          </Field>
          <Button variant="primary" disabled={!sender.trim() && !recipient.trim()} onClick={runTrace}>
            <MailSearch size={15} /> {t('mailTrace.run')}
          </Button>
          {tracePrereq === false && (
            <ErrorNote>
              <p>{t('mailTrace.prereqMissing')}</p>
              <Button variant="subtle" className="mt-2" disabled={readOnly || enabling} onClick={enableTrace}>
                <PlugZap size={15} /> {t('mailTrace.enable')}
              </Button>
            </ErrorNote>
          )}
          <div className="mt-2 border-t border-[var(--border)] pt-3">
            <p className="text-sm font-medium">{t('mailTrace.detailsTitle')}</p>
            <p className="mb-2 text-xs text-[var(--text-dim)]">{t('mailTrace.detailsHint')}</p>
            <Field label={t('mailTrace.traceId')}>
              <Input value={traceId} onChange={(e) => setTraceId(e.target.value)} placeholder="7e3b2b2e-…" autoComplete="off" />
            </Field>
            <Field label={t('mailTrace.recipient')}>
              <UpnInput value={traceRecipient} onChange={setTraceRecipient} placeholder="recipient@contoso.com" />
            </Field>
            <Button
              variant="subtle"
              disabled={!traceId.trim() || !traceRecipient.trim()}
              onClick={() => res.run(() => api.mailTrace.details(traceId.trim(), traceRecipient.trim()))}
            >
              <ListTree size={15} /> {t('mailTrace.details')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'signins', label: t('audit.tileSignIns'), hint: t('audit.hintSignIns'), icon: <LogIn size={16} />, variant: 'primary',
      panel: (
        <TaskForm>
          {period}
          {limit}
          <Button variant="primary" onClick={() => res.run(() => api.auditQuery.signIns('', days, false, top))}>
            <LogIn size={15} /> {t('common.run')}
          </Button>
          <Button variant="subtle" onClick={() => res.run(() => api.auditQuery.signIns('', days, true, top))}>
            <ShieldX size={15} /> {t('audit.failedOnly')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'directory', label: t('audit.tileDirectory'), hint: t('audit.hintDirectory'), icon: <FileClock size={16} />, variant: 'primary',
      note: <p>{t('audit.noteDirectory')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('audit.actor')} hint={t('audit.actorHint')}>
            <EntityPicker value={actor} onChange={setActor} load={loadUsers} placeholder={t('audit.anyActor')} />
          </Field>
          {period}
          {limit}
          <Button variant="primary" onClick={() => res.run(() => api.auditQuery.directory(actor, days, top))}>
            <FileClock size={15} /> {t('common.run')}
          </Button>
          {actor && (
            <Button variant="ghost" onClick={() => setActor('')}>{t('audit.clearActor')}</Button>
          )}
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      <TaskPage
        pageId="audit"
        title={t('nav.audit')}
        subtitle={t('audit.subtitle')}
        actions={actions}
        busy={res.loading || enabling}
        busyLabel={enabling ? t('mailTrace.enableTitle') : t('audit.querying')}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={useTraceRow} />}
      />
      {confirmElement}
    </>
  )
}
