import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Smartphone, Info, Eraser, Trash, Lock } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { EntityPicker } from '../components/EntityPicker'
import { loadIntuneDevices } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function IntunePage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [deviceId, setDeviceId] = useState('')
  const [keepEnroll, setKeepEnroll] = useState(false)
  const [keepUser, setKeepUser] = useState(false)

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    { id: 'list', label: 'Managed devices', icon: <Smartphone size={15} />, variant: 'primary', onClick: () => res.run(() => api.intune.devices(0)) },
    {
      id: 'device', label: t('devices.info'), icon: <Info size={15} />,
      panel: (
        <DrawerForm>
          <Field label="Device"><EntityPicker value={deviceId} onChange={setDeviceId} load={loadIntuneDevices} placeholder="Pick a device…" /></Field>
          <Button variant="subtle" disabled={!deviceId} onClick={() => res.run(() => api.intune.device(deviceId))}><Info size={15} /> {t('devices.info')}</Button>
        </DrawerForm>
      ),
    },
    {
      id: 'actions', label: 'Actions', icon: <Eraser size={15} />, variant: 'danger',
      panel: (
        <DrawerForm>
          <Field label="Device"><EntityPicker value={deviceId} onChange={setDeviceId} load={loadIntuneDevices} placeholder="Pick a device…" /></Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={keepEnroll} onChange={(e) => setKeepEnroll(e.target.checked)} /> keep enrollment data
          </label>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={keepUser} onChange={(e) => setKeepUser(e.target.checked)} /> keep user data
          </label>
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.wipe(deviceId, keepEnroll, keepUser, c), 'Wipe'))}><Eraser size={15} /> Wipe</Button>
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.retire(deviceId, c), 'Retire'))}><Trash size={15} /> Retire</Button>
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.remoteLock(deviceId, c), 'Remote lock'))}><Lock size={15} /> Remote lock</Button>
        </DrawerForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <ActionPage title={t('nav.intune')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setDeviceId} />} />
    </>
  )
}
