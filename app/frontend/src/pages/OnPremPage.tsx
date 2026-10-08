import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Server, Plug, LogOut, Trash2, Pencil, FolderOpen, ShieldAlert } from 'lucide-react'
import { catalogTile } from '../components/CatalogAction'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { Button, Field, Input, Select, Badge } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'

type Conn = { id: string; name: string; host: string; port: number; tls: string; baseDn: string; bindDn: string; caFile?: string; hasPassword?: boolean }
const EMPTY: Conn = { id: '', name: '', host: '', port: 0, tls: 'ldaps', baseDn: '', bindDn: '', caFile: '' }

// On-premises Active Directory: saved LDAP connections and the AD actions.
// Works without a tenant connection — hybrid admins use both side by side.
export function OnPremPage() {
  const { t } = useTranslation()
  const { toast, setCache } = useStore()
  const [conns, setConns] = useState<Conn[]>([])
  const [status, setStatus] = useState<{ connected: boolean; id?: string; name?: string; host?: string; secure?: boolean }>({ connected: false })
  const [form, setForm] = useState<Conn>(EMPTY)
  const [password, setPassword] = useState('')
  const [busy, setBusy] = useState(false)

  const load = () => {
    api.onprem.connections().then((l) => setConns((l ?? []) as Conn[])).catch(() => {})
    api.onprem.status().then((s) => setStatus((s as any) ?? { connected: false })).catch(() => {})
  }
  useEffect(load, [])
  // The catalog's availability depends on the directory connection.
  const refreshCatalog = () => setCache('catalog.rev', Date.now())

  const connect = async (id: string) => {
    setBusy(true)
    try {
      setStatus(((await api.onprem.connect(id)) as any) ?? { connected: false })
      refreshCatalog()
      toast('ok', t('onprem.connected'))
    } catch (e) { toast('err', errMessage(e)) } finally { setBusy(false) }
  }
  const disconnect = async () => {
    await api.onprem.disconnect()
    setStatus({ connected: false })
    refreshCatalog()
  }
  const save = async () => {
    try {
      await api.onprem.save({ ...form, port: Number(form.port) || 0 }, password)
      setForm(EMPTY); setPassword('')
      toast('ok', t('common.save'))
      load(); refreshCatalog()
    } catch (e) { toast('err', errMessage(e)) }
  }
  const remove = async (c: Conn) => {
    try { await api.onprem.remove(c.id); load(); refreshCatalog() } catch (e) { toast('err', errMessage(e)) }
  }
  const pickCA = async () => {
    try { const p = await api.onprem.pickCA(); if (p) setForm({ ...form, caFile: p }) } catch (e) { toast('err', errMessage(e)) }
  }

  const actions: TaskAction[] = [
    {
      id: 'connection', label: t('onprem.tileConnection'), hint: t('onprem.hintConnection'), icon: <Server size={16} />, variant: 'primary',
      panel: (
        <TaskForm>
          {status.connected ? (
            <div className="flex items-center justify-between rounded-lg border border-[var(--ok)]/30 bg-[var(--ok)]/10 px-3 py-2 text-sm">
              <span className="text-[var(--ok)]">{t('onprem.connectedTo', { name: status.name, host: status.host })}</span>
              <Button variant="ghost" onClick={disconnect} className="!px-2 !py-1"><LogOut size={14} /> {t('connect.disconnect')}</Button>
            </div>
          ) : <p className="text-sm text-[var(--text-faint)]">{t('onprem.notConnected')}</p>}

          <div className="flex flex-col gap-1.5">
            {conns.map((c) => (
              <div key={c.id} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2">
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium">{c.name}</div>
                  <div className="truncate text-xs text-[var(--text-faint)]">{c.host} · {c.baseDn}</div>
                </div>
                <Badge kind={c.tls === 'none' ? 'warn' : 'neutral'}>{t(`onprem.tls.${c.tls}`)}</Badge>
                <Button variant="primary" disabled={busy} onClick={() => connect(c.id)} className="!px-2 !py-1" title={t('onprem.connect')}><Plug size={14} /></Button>
                <Button variant="ghost" onClick={() => { setForm({ ...EMPTY, ...c }); setPassword('') }} className="!px-2 !py-1" title={t('onprem.edit')}><Pencil size={14} /></Button>
                <Button variant="ghost" onClick={() => remove(c)} className="!px-2 !py-1" title={t('common.delete')}><Trash2 size={14} /></Button>
              </div>
            ))}
          </div>

          <p className="mt-2 text-xs font-medium text-[var(--text-dim)]">{form.id ? t('onprem.editTitle') : t('onprem.addTitle')}</p>
          <Field label={t('onprem.name')}><Input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="CORP" /></Field>
          <div className="grid grid-cols-[1fr_6rem] gap-2">
            <Field label={t('onprem.host')}><Input value={form.host} onChange={(e) => setForm({ ...form, host: e.target.value })} placeholder="dc01.corp.example" /></Field>
            <Field label={t('onprem.port')}><Input value={form.port ? String(form.port) : ''} onChange={(e) => setForm({ ...form, port: Number(e.target.value.replace(/\D/g, '')) || 0 })} placeholder={form.tls === 'ldaps' ? '636' : '389'} /></Field>
          </div>
          <Field label={t('onprem.tlsMode')}>
            <Select value={form.tls} onChange={(e) => setForm({ ...form, tls: e.target.value })} className="w-full">
              {['ldaps', 'starttls', 'none'].map((m) => <option key={m} value={m}>{t(`onprem.tls.${m}`)}</option>)}
            </Select>
          </Field>
          {form.tls === 'none' && (
            <p className="flex items-start gap-1.5 text-xs text-[var(--warn)]"><ShieldAlert size={13} className="mt-0.5 shrink-0" /> {t('onprem.plainWarn')}</p>
          )}
          <Field label={t('onprem.baseDn')}><Input value={form.baseDn} onChange={(e) => setForm({ ...form, baseDn: e.target.value })} placeholder="DC=corp,DC=example" /></Field>
          <Field label={t('onprem.bindDn')} hint={t('onprem.bindHint')}><Input value={form.bindDn} onChange={(e) => setForm({ ...form, bindDn: e.target.value })} placeholder="svc-swissknife@corp.example" /></Field>
          <Field label={t('onprem.password')} hint={form.id ? t('onprem.passwordKept') : t('onprem.passwordHint')}>
            <Input type="password" autoComplete="new-password" value={password} onChange={(e) => setPassword(e.target.value)} />
          </Field>
          {form.tls !== 'none' && (
            <Field label={t('onprem.caFile')} hint={t('onprem.caHint')}>
              <div className="flex gap-2">
                <Input value={form.caFile || ''} onChange={(e) => setForm({ ...form, caFile: e.target.value })} className="min-w-0 flex-1" />
                <Button variant="subtle" onClick={pickCA}><FolderOpen size={14} /></Button>
              </div>
            </Field>
          )}
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" onClick={save} disabled={!form.host || !form.baseDn || !form.bindDn || (!form.id && !password)}>{t('common.save')}</Button>
            {form.id && <Button variant="subtle" onClick={() => { setForm(EMPTY); setPassword('') }}>{t('common.cancel')}</Button>}
          </div>
        </TaskForm>
      ),
    },
    catalogTile('ad.findUser'),
    catalogTile('ad.userState'),
    catalogTile('ad.unlock'),
    catalogTile('ad.resetPassword'),
    catalogTile('ad.groupMembership'),
  ]

  return (
    <TaskPage pageId="onprem" title={t('nav.onprem')} subtitle={t('onprem.subtitle')} actions={actions}
      hasResult={false} result={null} />
  )
}
