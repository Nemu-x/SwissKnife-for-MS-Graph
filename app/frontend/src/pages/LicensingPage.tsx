import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { List, UserCheck, Plus, Minus } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'

export function LicensingPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const res = useAsync<GraphObject[] | GraphObject>()
  const [target, setTarget] = useState('')
  const [skuId, setSkuId] = useState('')

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  return (
    <Page title={t('nav.licensing')}>
      <TwoPane
        controls={
          <>
            <Card title={t('nav.licensing')}>
              <div className="flex flex-col gap-2">
                <Button variant="primary" onClick={() => res.run(() => api.licensing.skus())}>
                  <List size={15} /> Tenant SKUs
                </Button>
                <Field label={t('common.user')}>
                  <Input value={target} onChange={(e) => setTarget(e.target.value)} placeholder="user@contoso.com" />
                </Field>
                <Button variant="subtle" disabled={!target} onClick={() => res.run(() => api.users.licenseDetails(target))}>
                  <UserCheck size={15} /> User licenses
                </Button>
              </div>
            </Card>

            <Card title="Assign / remove">
              <div className="flex flex-col gap-3">
                <Field label="SKU ID (GUID)">
                  <Input value={skuId} onChange={(e) => setSkuId(e.target.value)} placeholder="00000000-0000-…" />
                </Field>
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" disabled={readOnly || !target || !skuId}
                    onClick={() => doWrite(() => api.licensing.assign(target, [skuId], []), t('common.add'))}>
                    <Plus size={15} /> {t('common.add')}
                  </Button>
                  <Button variant="subtle" disabled={readOnly || !target || !skuId}
                    onClick={() => doWrite(() => api.licensing.assign(target, [], [skuId]), t('common.remove'))}>
                    <Minus size={15} /> {t('common.remove')}
                  </Button>
                </div>
              </div>
            </Card>
          </>
        }
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </Page>
  )
}
