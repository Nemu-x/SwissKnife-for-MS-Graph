import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, Info, Power, PowerOff, Trash2, KeyRound, Eye } from 'lucide-react'
import { ActionPage, DrawerForm, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function DevicesPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [search, setSearch] = useState('')
  const [deviceId, setDeviceId] = useState('')
  const [keyId, setKeyId] = useState('')

  const listDevices = () => res.run(() => api.devices.list(search, 0))
  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: Action[] = [
    { id: 'search', label: t('common.search'), icon: <Search size={15} />, variant: 'primary', onClick: listDevices },
    {
      id: 'device', label: t('devices.info'), icon: <Info size={15} />,
      panel: (
        <DrawerForm>
          <Field label={t('devices.deviceId')}><Input value={deviceId} onChange={(e) => setDeviceId(e.target.value)} /></Field>
          <Button variant="subtle" disabled={!deviceId} onClick={() => res.run(() => api.devices.get(deviceId))}><Info size={15} /> {t('devices.info')}</Button>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={readOnly || !deviceId} onClick={() => doWrite(() => api.devices.enable(deviceId), t('devices.enable'))}><Power size={15} /> {t('devices.enable')}</Button>
            <Button variant="subtle" disabled={readOnly || !deviceId} onClick={() => doWrite(() => api.devices.disable(deviceId), t('devices.disable'))}><PowerOff size={15} /> {t('devices.disable')}</Button>
          </div>
          <Button variant="danger" disabled={readOnly || !deviceId}
            onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.devices.delete(deviceId, c), t('devices.delete')))}>
            <Trash2 size={15} /> {t('devices.delete')}
          </Button>
        </DrawerForm>
      ),
    },
    {
      id: 'bitlocker', label: t('devices.bitlocker'), icon: <KeyRound size={15} />,
      panel: (
        <DrawerForm>
          <Button variant="subtle" onClick={() => res.run(() => api.devices.bitlockerKeys(0))}><KeyRound size={15} /> {t('devices.listKeys')}</Button>
          <Field label={t('devices.keyId')}><Input value={keyId} onChange={(e) => setKeyId(e.target.value)} /></Field>
          <Button variant="subtle" disabled={!keyId} onClick={() => res.run(() => api.devices.bitlockerKey(keyId) as any)}><Eye size={15} /> {t('devices.revealKey')}</Button>
        </DrawerForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      <ActionPage
        title={t('devices.title')}
        search={{ value: search, onChange: setSearch, onSubmit: listDevices, placeholder: t('common.search') }}
        actions={actions}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </>
  )
}
