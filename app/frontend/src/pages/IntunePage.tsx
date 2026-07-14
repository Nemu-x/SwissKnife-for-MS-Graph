import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Smartphone, Info, Eraser, Trash, Lock } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input } from '../components/ui'
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

  return (
    <Page title={t('nav.intune')}>
      {confirmElement}
      <TwoPane
        controls={
          <>
            <Card title={t('nav.intune')}>
              <div className="flex flex-col gap-2">
                <Button variant="primary" onClick={() => res.run(() => api.intune.devices(0))}>
                  <Smartphone size={15} /> Managed devices
                </Button>
                <Field label="Device ID"><Input value={deviceId} onChange={(e) => setDeviceId(e.target.value)} /></Field>
                <Button variant="subtle" disabled={!deviceId} onClick={() => res.run(() => api.intune.device(deviceId))}>
                  <Info size={15} /> Device info
                </Button>
              </div>
            </Card>

            <Card title="Actions (destructive)">
              <div className="flex flex-col gap-2">
                <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                  <input type="checkbox" checked={keepEnroll} onChange={(e) => setKeepEnroll(e.target.checked)} /> keep enrollment data
                </label>
                <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                  <input type="checkbox" checked={keepUser} onChange={(e) => setKeepUser(e.target.checked)} /> keep user data
                </label>
                <Button variant="danger" disabled={readOnly || !deviceId}
                  onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.wipe(deviceId, keepEnroll, keepUser, c), 'Wipe'))}>
                  <Eraser size={15} /> Wipe
                </Button>
                <Button variant="danger" disabled={readOnly || !deviceId}
                  onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.retire(deviceId, c), 'Retire'))}>
                  <Trash size={15} /> Retire
                </Button>
                <Button variant="danger" disabled={readOnly || !deviceId}
                  onClick={() => askConfirm(deviceId, (c) => doWrite(() => api.intune.remoteLock(deviceId, c), 'Remote lock'))}>
                  <Lock size={15} /> Remote lock
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
