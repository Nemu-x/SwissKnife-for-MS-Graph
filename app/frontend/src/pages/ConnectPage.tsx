import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Plug, Trash2, LogOut } from 'lucide-react'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { Page } from '../components/Layout'
import { Button, Card, Field, Input, Select, Badge, ErrorNote, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { api, errMessage, type Profile } from '../lib/api'

export function ConnectPage() {
  const { t } = useTranslation()
  const { status, setStatus, refreshStatus, connected, toast, loadDomains } = useStore()

  const [profiles, setProfiles] = useState<Profile[]>([])
  const [form, setForm] = useState({
    id: '', name: '', tenantId: '', clientId: '', secret: '', authMode: 'client_secret', remember: false,
  })
  const [wantDomains, setWantDomains] = useState(localStorage.getItem('loadDomains') === 'true')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [device, setDevice] = useState<{ url: string; code: string } | null>(null)

  const loadProfiles = () => api.connect.profiles().then(setProfiles).catch(() => {})
  useEffect(() => { loadProfiles() }, [])

  useEffect(() => {
    const off = EventsOn('auth:deviceCode', (d: any) => setDevice({ url: d.url, code: d.code }))
    return () => off()
  }, [])

  const selectProfile = (p: Profile) =>
    setForm({ id: p.id, name: p.name, tenantId: p.tenantId, clientId: p.clientId, secret: '', authMode: p.authMode, remember: true })

  const connectProfile = async (p: Profile) => {
    setBusy(true); setError(null); setDevice(null)
    try {
      const s = await api.connect.connect({ profileId: p.id })
      setStatus(s); toast('ok', t('common.connectedAs', { name: s.profileName }))
      if (wantDomains) loadDomains()
    } catch (e) { setError(errMessage(e)) } finally { setBusy(false); setDevice(null) }
  }

  const connectForm = async () => {
    setBusy(true); setError(null); setDevice(null)
    try {
      const s = await api.connect.connect({
        tenantId: form.tenantId, clientId: form.clientId, secret: form.secret,
        authMode: form.authMode, rememberAs: form.remember ? form.name : '',
      })
      setStatus(s); toast('ok', t('common.connectedAs', { name: s.profileName })); loadProfiles()
      if (wantDomains) loadDomains()
    } catch (e) { setError(errMessage(e)) } finally { setBusy(false); setDevice(null) }
  }

  const saveProfile = async () => {
    try {
      await api.connect.saveProfile(
        { id: form.id, name: form.name, tenantId: form.tenantId, clientId: form.clientId, authMode: form.authMode },
        form.secret,
      )
      toast('ok', t('common.save')); loadProfiles()
    } catch (e) { toast('err', errMessage(e)) }
  }

  const deleteProfile = async (p: Profile) => {
    try { await api.connect.deleteProfile(p.id); loadProfiles() } catch (e) { toast('err', errMessage(e)) }
  }

  const disconnect = async () => { await api.connect.disconnect(); refreshStatus() }

  return (
    <Page title={t('connect.title')}>
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <Card title={t('connect.profiles')}>
          {profiles.length === 0 && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          <div className="flex flex-col gap-2">
            {profiles.map((p) => (
              <div key={p.id} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2">
                <div className="min-w-0 flex-1 cursor-pointer" onClick={() => selectProfile(p)}>
                  <div className="truncate text-sm font-medium">{p.name}</div>
                  <div className="truncate text-xs text-[var(--text-faint)]">{p.tenantId}</div>
                </div>
                <Badge kind="neutral">{p.authMode === 'device_code' ? 'device' : 'app'}</Badge>
                <Button variant="primary" onClick={() => connectProfile(p)} disabled={busy} className="!px-2 !py-1">
                  <Plug size={14} />
                </Button>
                <Button variant="ghost" onClick={() => deleteProfile(p)} className="!px-2 !py-1">
                  <Trash2 size={14} />
                </Button>
              </div>
            ))}
          </div>
          {connected && (
            <div className="mt-4 flex items-center justify-between rounded-lg border border-[var(--ok)]/30 bg-[var(--ok)]/10 px-3 py-2">
              <span className="text-sm text-[var(--ok)]">{t('common.connectedAs', { name: status?.profileName })}</span>
              <Button variant="ghost" onClick={disconnect} className="!px-2 !py-1">
                <LogOut size={14} /> {t('connect.disconnect')}
              </Button>
            </div>
          )}
        </Card>

        <Card title={form.id ? form.name : t('connect.newProfile')}>
          <div className="flex flex-col gap-3">
            <Field label={t('connect.name')}>
              <Input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </Field>
            <Field label={t('connect.authMode')}>
              <Select value={form.authMode} onChange={(e) => setForm({ ...form, authMode: e.target.value })} className="w-full">
                <option value="client_secret">{t('connect.clientSecret')}</option>
                <option value="device_code">{t('connect.deviceCode')}</option>
              </Select>
            </Field>
            <Field label={t('connect.tenantId')}>
              <Input value={form.tenantId} onChange={(e) => setForm({ ...form, tenantId: e.target.value })} />
            </Field>
            <Field label={t('connect.clientId')}>
              <Input value={form.clientId} onChange={(e) => setForm({ ...form, clientId: e.target.value })} />
            </Field>
            {form.authMode === 'client_secret' && (
              <Field label={t('connect.secret')} hint={form.id ? t('connect.secretKept') : t('connect.rememberHint')}>
                <Input type="password" value={form.secret} onChange={(e) => setForm({ ...form, secret: e.target.value })} />
              </Field>
            )}
            <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
              <input type="checkbox" checked={form.remember} onChange={(e) => setForm({ ...form, remember: e.target.checked })} />
              {t('connect.remember')}
            </label>
            <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
              <input type="checkbox" checked={wantDomains}
                onChange={(e) => { setWantDomains(e.target.checked); localStorage.setItem('loadDomains', String(e.target.checked)) }} />
              {t('connect.loadDomains')}
            </label>

            {device && (
              <div className="rounded-lg border border-[var(--accent)]/40 bg-[var(--accent)]/10 p-3 text-sm">
                <div className="font-medium">{t('connect.deviceCodeTitle')}</div>
                <div className="mt-1 text-[var(--text-dim)]">{t('connect.deviceCodeBody', { url: device.url })}</div>
                <div className="mt-2 select-all rounded bg-[var(--bg)] px-2 py-1 text-center font-mono text-lg tracking-widest">
                  {device.code}
                </div>
              </div>
            )}
            {error && <ErrorNote>{error}</ErrorNote>}

            <div className="flex gap-2">
              <Button variant="primary" onClick={connectForm} disabled={busy} className="flex-1">
                {busy ? <Spinner /> : <Plug size={15} />}
                {busy ? t('connect.connecting') : t('connect.connect')}
              </Button>
              <Button variant="ghost" onClick={saveProfile}>{t('common.save')}</Button>
            </div>
          </div>
        </Card>
      </div>
    </Page>
  )
}
