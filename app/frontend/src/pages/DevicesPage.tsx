import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Info, Power, PowerOff, Trash2, KeyRound, Eye } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadDevices } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useTaskStatus } from '../lib/useTaskStatus'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, type GraphObject } from '../lib/api'

export function DevicesPage() {
  const { t } = useTranslation()
  const { readOnly } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [deviceId, setDeviceId] = useState('')
  const [keyId, setKeyId] = useState('')
  const { status, busy: writing, doWrite } = useTaskStatus()

  const listDevices = () => res.run(() => api.devices.list(search, 0))

  const deviceField = (
    <Field label={t('devices.device')}>
      <EntityPicker value={deviceId} onChange={setDeviceId} load={loadDevices} placeholder={t('devices.pickDevice')} />
    </Field>
  )

  const actions: TaskAction[] = [
    { id: 'list', label: t('devices.tileList'), hint: t('devices.hintList'), icon: <Search size={16} />, variant: 'primary', onClick: listDevices },
    {
      id: 'info', label: t('devices.tileInfo'), hint: t('devices.hintInfo'), icon: <Info size={16} />, write: true,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="primary" disabled={!deviceId} onClick={() => res.run(() => api.devices.get(deviceId))}>
            <Info size={15} /> {t('devices.info')}
          </Button>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !deviceId} onClick={() => doWrite('info', () => api.devices.enable(deviceId), t('devices.enable'))}>
              <Power size={15} /> {t('devices.enable')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !deviceId} onClick={() => doWrite('info', () => api.devices.disable(deviceId), t('devices.disable'))}>
              <PowerOff size={15} /> {t('devices.disable')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'delete', label: t('devices.tileDelete'), hint: t('devices.hintDelete'), icon: <Trash2 size={16} />, variant: 'danger', write: true,
      note: <p className="text-[var(--danger)]">{t('devices.noteDelete')}</p>,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite('delete', () => api.devices.delete(deviceId, c), t('devices.delete')))}>
            <Trash2 size={15} /> {t('devices.delete')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'bitlocker', label: t('devices.tileBitlocker'), hint: t('devices.hintBitlocker'), icon: <KeyRound size={16} />,
      note: <p>{t('devices.noteBitlocker')}</p>,
      panel: (
        <TaskForm>
          {deviceField}
          <Button variant="primary" disabled={!deviceId} onClick={() => res.run(() => api.devices.bitlockerKeysForDevice(deviceId))}>
            <KeyRound size={15} /> {t('devices.keysOfDevice')}
          </Button>
          <Field label={t('devices.keyId')} hint={t('devices.keyIdHint')}>
            <Input value={keyId} onChange={(e) => setKeyId(e.target.value)} />
          </Field>
          <Button variant="primary" disabled={!keyId} onClick={() => res.run(() => api.devices.bitlockerKey(keyId) as any)}>
            <Eye size={15} /> {t('devices.revealKey')}
          </Button>
          <div className="my-1 border-t border-[var(--border)]" />
          <Button variant="subtle" onClick={() => res.run(() => api.devices.bitlockerKeys(0))}>
            {t('devices.listKeys')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="devices"
        title={t('devices.title')}
        subtitle={t('devices.subtitle')}
        search={{ value: search, onChange: setSearch, onSubmit: listDevices, placeholder: t('common.search') }}
        actions={actions}
        status={status}
        busy={res.loading || writing}
        onClearResult={res.reset}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error}
          onUseId={(id) => { setDeviceId(id); setKeyId(id) }} />}
      />
    </>
  )
}
