import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import {
  Search, UserCog, Ban, CheckCircle, KeyRound, LogOut, UserPlus, Trash2,
  UserSquare, MapPin, ShieldAlert, ListRestart, RotateCcw,
} from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card, Field, Input, Textarea } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
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
  const [create, setCreate] = useState({ name: '', upn: '', nick: '', pw: '', loc: '' })
  const [mgr, setMgr] = useState('')
  const [loc, setLoc] = useState('')
  const [restoreId, setRestoreId] = useState('')
  const [patch, setPatch] = useState('{\n  "jobTitle": "",\n  "department": ""\n}')

  const listUsers = () => res.run(() => api.users.list(search, 0))
  const snapshot = () => target && res.run(() => api.users.snapshot(target) as any)

  const doWrite = async (fn: () => Promise<any>, ok: string) => {
    try { await fn(); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
  }
  const doWriteShow = async (fn: () => Promise<any>, ok: string) => {
    try { res.setData(await fn()); toast('ok', ok) } catch (e) { toast('err', errMessage(e)) }
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
                  <UpnInput value={target} onChange={setTarget} />
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

            <Card title={t('users.security')}>
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
                <div className="grid grid-cols-2 gap-2">
                  <Button variant="subtle" disabled={!target} onClick={() => target && res.run(() => api.authMethods.list(target))}>
                    <ShieldAlert size={15} /> {t('users.authMethods')}
                  </Button>
                  <Button variant="danger" disabled={readOnly || !target}
                    onClick={() => askConfirm(target, (c) => doWriteShow(() => api.authMethods.resetMFA(target, c), t('users.resetMfa')))}>
                    <RotateCcw size={15} /> {t('users.resetMfa')}
                  </Button>
                </div>
              </div>
            </Card>

            <Card title={t('users.directory')}>
              <div className="flex flex-col gap-2">
                <Button variant="subtle" disabled={!target} onClick={() => res.run(() => api.users.getManager(target))}>
                  <UserSquare size={15} /> {t('users.getManager')}
                </Button>
                <Field label={t('users.manager')}>
                  <UpnInput value={mgr} onChange={setMgr} placeholder="manager@contoso.com" />
                </Field>
                <Button variant="subtle" disabled={readOnly || !target || !mgr}
                  onClick={() => doWrite(() => api.users.setManager(target, mgr), t('users.setManager'))}>
                  {t('users.setManager')}
                </Button>
                <div className="flex gap-2">
                  <Input value={loc} onChange={(e) => setLoc(e.target.value)} placeholder={t('users.usageLocation')} />
                  <Button variant="subtle" disabled={readOnly || !target || !loc}
                    onClick={() => doWrite(() => api.users.setUsageLocation(target, loc), t('users.setUsageLocation'))}>
                    <MapPin size={15} />
                  </Button>
                </div>
                <Field label={t('users.updateFields')}>
                  <Textarea rows={4} value={patch} onChange={(e) => setPatch(e.target.value)} />
                </Field>
                <Button variant="subtle" disabled={readOnly || !target}
                  onClick={() => doWrite(() => api.users.update(target, patch), t('users.update'))}>
                  {t('users.update')}
                </Button>
              </div>
            </Card>

            <Card title={t('users.lifecycle')}>
              <div className="flex flex-col gap-2">
                <Input placeholder={t('users.displayName')} value={create.name} onChange={(e) => setCreate({ ...create, name: e.target.value })} />
                <Input placeholder="user@contoso.com" value={create.upn} onChange={(e) => setCreate({ ...create, upn: e.target.value })} />
                <Input placeholder={t('users.mailNickname')} value={create.nick} onChange={(e) => setCreate({ ...create, nick: e.target.value })} />
                <Input type="password" placeholder={t('users.newPassword')} value={create.pw} onChange={(e) => setCreate({ ...create, pw: e.target.value })} />
                <Input placeholder={t('users.usageLocation')} value={create.loc} onChange={(e) => setCreate({ ...create, loc: e.target.value })} />
                <Button variant="primary" disabled={readOnly || !create.name || !create.upn || !create.nick || !create.pw}
                  onClick={() => doWriteShow(() => api.users.create(create.name, create.upn, create.nick, create.pw, true, create.loc), t('users.createUser'))}>
                  <UserPlus size={15} /> {t('users.createUser')}
                </Button>
                <Button variant="danger" disabled={readOnly || !target}
                  onClick={() => askConfirm(target, (c) => doWrite(() => api.users.delete(target, c), t('users.deleteUser')))}>
                  <Trash2 size={15} /> {t('users.deleteUser')}
                </Button>
              </div>
            </Card>

            <Card title={t('users.deleted')}>
              <div className="flex flex-col gap-2">
                <Button variant="subtle" onClick={() => res.run(() => api.users.listDeleted(0))}>
                  <ListRestart size={15} /> {t('users.listDeleted')}
                </Button>
                <Field label={t('users.objectId')}>
                  <Input value={restoreId} placeholder="object id" onChange={(e) => setRestoreId(e.target.value)} />
                </Field>
                <Button variant="subtle" disabled={readOnly || !restoreId}
                  onClick={() => doWriteShow(() => api.users.restoreDeleted(restoreId), t('users.restore'))}>
                  <RotateCcw size={15} /> {t('users.restore')}
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
