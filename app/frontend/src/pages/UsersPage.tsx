import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search, UserCog, Ban, CheckCircle, KeyRound, LogOut } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import type { GraphObject } from '../lib/api'

export function UsersPage() {
  const { t } = useTranslation()
  const { readOnly, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()

  const [search, setSearch] = useState('')
  const [target, setTarget] = useState('')
  const [pw, setPw] = useState('')
  const [force, setForce] = useState(true)

  const listUsers = () => res.run(() => api.users.list(search, 0))
  const snapshot = () => target && res.run(() => api.users.snapshot(target) as any)

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }

  return (
    <Page title={t('nav.users')}>
      {confirmElement}
      <TwoPane
        controls={
          <>
            <Card title={t('users.listTitle')}>
              <div className="flex gap-2">
                <Input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={t('common.search')} />
                <Button variant="primary" onClick={listUsers}><Search size={15} /></Button>
              </div>
            </Card>

            <Card title={t('nav.users')}>
              <div className="flex flex-col gap-3">
                <Field label={t('common.user')}>
                  <Input value={target} onChange={(e) => setTarget(e.target.value)} placeholder="user@contoso.com" />
                </Field>
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" onClick={snapshot}><UserCog size={15} /> {t('users.snapshot')}</Button>
                  <Button variant="subtle" disabled={readOnly} onClick={() => doWrite(() => api.users.block(target), t('users.block'))}>
                    <Ban size={15} /> {t('users.block')}
                  </Button>
                  <Button variant="subtle" disabled={readOnly} onClick={() => doWrite(() => api.users.unblock(target), t('users.unblock'))}>
                    <CheckCircle size={15} /> {t('users.unblock')}
                  </Button>
                  <Button variant="subtle" disabled={readOnly}
                    onClick={() => askConfirm(target, (c) => doWrite(() => api.users.revokeSessions(target, c), t('users.revokeSessions')))}>
                    <LogOut size={15} /> {t('users.revokeSessions')}
                  </Button>
                </div>
              </div>
            </Card>

            <Card title={t('users.resetPassword')}>
              <div className="flex flex-col gap-3">
                <Field label={t('users.newPassword')}>
                  <Input type="password" value={pw} onChange={(e) => setPw(e.target.value)} />
                </Field>
                <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                  <input type="checkbox" checked={force} onChange={(e) => setForce(e.target.checked)} />
                  {t('users.forceChange')}
                </label>
                <Button variant="danger" disabled={readOnly || !pw || !target}
                  onClick={() => askConfirm(target, (c) => doWrite(() => api.users.resetPassword(target, pw, force, c), t('users.resetPassword')))}>
                  <KeyRound size={15} /> {t('users.resetPassword')}
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
