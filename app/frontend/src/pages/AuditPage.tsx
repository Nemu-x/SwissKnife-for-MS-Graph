import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { LogIn, FileClock, ShieldX } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Select } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { api, type GraphObject } from '../lib/api'

const PERIODS = [1, 7, 30]

export function AuditPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[]>()
  const [upn, setUpn] = useState('')
  const [days, setDays] = useState(7)
  const [top, setTop] = useState(50)
  const [actor, setActor] = useState('')

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
    <TaskPage
      pageId="audit"
      title={t('nav.audit')}
      subtitle={t('audit.subtitle')}
      actions={actions}
      busy={res.loading}
      busyLabel={t('audit.querying')}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
