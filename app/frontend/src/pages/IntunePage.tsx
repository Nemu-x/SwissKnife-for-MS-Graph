import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Smartphone, Info, Eraser, Trash, Lock } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadIntuneDevices } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function IntunePage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [deviceId, setDeviceId] = useState('')
  const [keepEnroll, setKeepEnroll] = useState(false)
  const [keepUser, setKeepUser] = useState(false)
  const { status, busy: writing, doWrite } = useTaskStatus()

  const deviceField = (
    <Field label={t('intune.device')}>
      <EntityPicker value={deviceId} onChange={setDeviceId} load={loadIntuneDevices} placeholder={t('intune.pickDevice')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'list', label: t('intune.tileList'), hint: t('intune.hintList'), icon: <Smartphone size={16} />, variant: 'primary',
      onClick: () => res.run(() => api.intune.devices(0)),
    },
    {
      id: 'info', label: t('intune.tileInfo'), hint: t('intune.hintInfo'), icon: <Info size={16} />,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="primary" disabled={!deviceId} onClick={() => res.run(() => api.intune.device(deviceId))}>
            <Info size={15} /> {t('devices.info')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'lock', label: t('intune.tileLock'), hint: t('intune.hintLock'), icon: <Lock size={16} />, write: true,
      note: <p>{t('intune.noteLock')}</p>,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite('lock', () => api.intune.remoteLock(deviceId, c), t('intune.remoteLock')))}>
            <Lock size={15} /> {t('intune.remoteLock')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'retire', label: t('intune.tileRetire'), hint: t('intune.hintRetire'), icon: <Trash size={16} />, write: true,
      note: <p>{t('intune.noteRetire')}</p>,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite('retire', () => api.intune.retire(deviceId, c), t('intune.retire')))}>
            <Trash size={15} /> {t('intune.retire')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'wipe', label: t('intune.tileWipe'), hint: t('intune.hintWipe'), icon: <Eraser size={16} />, variant: 'danger', write: true,
      note: <p className="text-[var(--danger)]">{t('intune.noteWipe')}</p>,
      panel: (
        <TaskForm>
          {deviceField}
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={keepEnroll} onChange={(e) => setKeepEnroll(e.target.checked)} /> {t('intune.keepEnrollment')}
          </label>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={keepUser} onChange={(e) => setKeepUser(e.target.checked)} /> {t('intune.keepUserData')}
          </label>
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite('wipe', () => api.intune.wipe(deviceId, keepEnroll, keepUser, c), t('intune.wipe')))}>
            <Eraser size={15} /> {t('intune.wipe')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="intune"
        title={t('nav.intune')}
        subtitle={t('intune.subtitle')}
        actions={actions}
        status={status}
        busy={res.loading || writing}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setDeviceId} />}
      />
    </>
  )
}
