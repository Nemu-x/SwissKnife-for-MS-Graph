import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Moon, Sun, Lock, RefreshCw, Download, Heart, Copy } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Select, Button, Badge, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { ACCENT_PRESETS, isHex } from '../lib/color'
import { setLanguage } from '../i18n'
import { api, errMessage } from '../lib/api'
import { Version } from '../../wailsjs/go/main/App'
import { EventsOn } from '../../wailsjs/runtime/runtime'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

// Crypto donation wallets shown in the Support card (copy-to-clipboard).
const SUPPORT_WALLETS: { asset: string; address: string }[] = [
  { asset: 'USDT (TRC20)', address: 'TPACN1kJRm2FnFF1cSqYtBnJwAmZ3qGMni' },
  { asset: 'USDT (Polygon)', address: '0xD9333e859Fb74D885d22E27568589de61E4433b5' },
  { asset: 'BTC', address: 'bc1qkkcgpqym967k2x73al6f7fpvkx52q4rzkut3we' },
  { asset: 'ETH', address: '0xD9333e859Fb74D885d22E27568589de61E4433b5' },
]

export function SettingsPage() {
  const { t, i18n } = useTranslation()
  const { theme, toggleTheme, accent, setAccent, safeMode, setSafeMode, hideUnavailable, setHideUnavailable, checkAccess, readOnly, setStatus, connected, toast } = useStore()
  const [checkingAccess, setCheckingAccess] = useState(false)
  const [version, setVersion] = useState('')
  const [update, setUpdate] = useState<services.UpdateInfo | null>(null)
  const [checking, setChecking] = useState(false)
  const [updating, setUpdating] = useState(false)
  const [updateProgress, setUpdateProgress] = useState('')

  useEffect(() => { Version().then(setVersion).catch(() => {}) }, [])

  // In-app update: download the installer with progress, then run it silently
  // (the app exits; the installer replaces files and relaunches).
  useEffect(() => EventsOn('update:progress', (d: any) => {
    const pct = d.total > 0 ? Math.round((d.done / d.total) * 100) : 0
    setUpdateProgress(`${humanBytes(d.done)} / ${humanBytes(d.total)} — ${pct}%`)
  }), [])

  // Teams webhook notifications (playbook summaries → channel card).
  const [notify, setNotify] = useState({ webhookUrl: '', notifyPlaybooks: false })
  const [notifyBusy, setNotifyBusy] = useState(false)
  useEffect(() => { api.notify.get().then((c) => setNotify({ webhookUrl: c.webhookUrl || '', notifyPlaybooks: !!c.notifyPlaybooks })).catch(() => {}) }, [])
  const saveNotify = async () => {
    setNotifyBusy(true)
    try { await api.notify.set(notify as any); toast('ok', t('settings.teamsSaved')) }
    catch (e) { toast('err', errMessage(e)) } finally { setNotifyBusy(false) }
  }
  const testNotify = async () => {
    setNotifyBusy(true)
    try { await api.notify.set(notify as any); await api.notify.test(); toast('ok', t('settings.teamsTestOk')) }
    catch (e) { toast('err', errMessage(e)) } finally { setNotifyBusy(false) }
  }

  const updateNow = async () => {
    if (!update?.assetUrl) return
    setUpdating(true); setUpdateProgress('')
    try {
      const path = await api.update.download(update.assetUrl, update.assetName || 'installer.exe', update.assetSize || 0)
      setUpdateProgress(t('settings.updateInstalling'))
      await api.update.apply(path)
      // the app quits moments later; nothing else to do
    } catch (e) {
      toast('err', errMessage(e))
      setUpdating(false)
    }
  }

  const toggleReadOnly = async () => setStatus(await api.connect.setReadOnly(!readOnly))

  const checkUpdates = async () => {
    setChecking(true)
    try { setUpdate(await api.update.check()) } catch (e) { toast('err', errMessage(e)) } finally { setChecking(false) }
  }

  return (
    <Page title={t('settings.title')}>
      <div className="grid max-w-2xl grid-cols-1 gap-4">
        <Card title={t('settings.language')}>
          <Select value={i18n.language.startsWith('ru') ? 'ru' : 'en'}
            onChange={(e) => setLanguage(e.target.value as 'en' | 'ru')} className="w-48">
            <option value="en">English</option>
            <option value="ru">Русский</option>
          </Select>
        </Card>

        <Card title={t('settings.theme')}>
          <Button variant="subtle" onClick={toggleTheme}>
            {theme === 'dark' ? <Moon size={15} /> : <Sun size={15} />}
            {theme === 'dark' ? t('settings.dark') : t('settings.light')}
          </Button>
        </Card>

        <Card title={t('settings.accent')}>
          <div className="flex flex-col gap-3">
            <div>
              <p className="mb-1.5 text-xs text-[var(--text-dim)]">{t('settings.accentPresets')}</p>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setAccent('')}
                  className={`rounded-full border px-3 py-1 text-xs transition-colors ${
                    accent === '' ? 'border-[var(--accent)] text-[var(--accent)]' : 'border-[var(--border)] text-[var(--text-dim)] hover:bg-[var(--bg-elev-2)]'
                  }`}
                >
                  {t('settings.accentAuto')}
                </button>
                {ACCENT_PRESETS.map((p) => (
                  <button
                    key={p.hex}
                    title={p.name}
                    onClick={() => setAccent(p.hex)}
                    className={`h-8 w-8 rounded-full border-2 transition-transform hover:scale-110 ${
                      accent.toLowerCase() === p.hex.toLowerCase() ? 'border-[var(--text)]' : 'border-transparent'
                    }`}
                    style={{ background: p.hex }}
                  />
                ))}
              </div>
            </div>
            <div className="flex items-center gap-2">
              <span className="text-xs text-[var(--text-dim)]">{t('settings.accentCustom')}</span>
              <input
                type="color"
                value={isHex(accent) ? accent : '#6366f1'}
                onChange={(e) => setAccent(e.target.value)}
                className="h-8 w-10 cursor-pointer rounded border border-[var(--border)] bg-transparent"
              />
              <input
                value={accent}
                onChange={(e) => setAccent(e.target.value)}
                placeholder="#6366f1"
                className="w-28 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-2 py-1 font-mono text-sm outline-none focus:border-[var(--accent)]"
              />
              <span className="h-6 w-6 rounded" style={{ background: isHex(accent) ? accent : 'transparent' }} />
            </div>
          </div>
        </Card>

        <Card title={t('safety.readOnly')}>
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" checked={safeMode} onChange={(e) => setSafeMode(e.target.checked)} />
            {t('settings.safeMode')}
          </label>
          {connected && (
            <Button variant={readOnly ? 'danger' : 'subtle'} onClick={toggleReadOnly} className="mt-3">
              <Lock size={15} /> {readOnly ? `${t('safety.readOnly')}: ON` : `${t('safety.readOnly')}: OFF`}
            </Button>
          )}
        </Card>

        <Card title={t('settings.access')}>
          <p className="text-xs text-[var(--text-faint)]">{t('settings.accessHint')}</p>
          <label className="mt-3 flex items-center gap-2 text-sm">
            <input type="checkbox" checked={hideUnavailable} onChange={(e) => setHideUnavailable(e.target.checked)} />
            {t('settings.hideUnavailable')}
          </label>
          {connected && (
            <Button variant="subtle" className="mt-3" disabled={checkingAccess}
              onClick={async () => { setCheckingAccess(true); await checkAccess(); setCheckingAccess(false); toast('ok', t('settings.checkAccess')) }}>
              {checkingAccess ? <Spinner /> : <RefreshCw size={15} />} {t('settings.checkAccess')}
            </Button>
          )}
        </Card>

        <Card title={t('settings.teams')}>
          <p className="text-xs text-[var(--text-faint)]">{t('settings.teamsHint')}</p>
          <div className="mt-3 flex flex-col gap-2">
            <input
              value={notify.webhookUrl}
              onChange={(e) => setNotify({ ...notify, webhookUrl: e.target.value })}
              placeholder="https://…logic.azure.com/workflows/…"
              className="w-full rounded-lg border border-[var(--border)] bg-[var(--bg)] px-2.5 py-1.5 font-mono text-xs outline-none focus:border-[var(--accent)]"
            />
            <label className="flex items-center gap-2 text-sm">
              <input type="checkbox" checked={notify.notifyPlaybooks}
                onChange={(e) => setNotify({ ...notify, notifyPlaybooks: e.target.checked })} />
              {t('settings.teamsPlaybooks')}
            </label>
            <div className="flex gap-2">
              <Button variant="subtle" disabled={notifyBusy} onClick={saveNotify}>
                {notifyBusy ? <Spinner /> : null} {t('common.save')}
              </Button>
              <Button variant="subtle" disabled={notifyBusy || !notify.webhookUrl} onClick={testNotify}>
                {t('settings.teamsTest')}
              </Button>
            </div>
          </div>
        </Card>

        <Card title={t('settings.support')}>
          <p className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <Heart size={15} className="shrink-0 text-[var(--danger)]" /> {t('settings.supportHint')}
          </p>
          <div className="mt-3 flex flex-col gap-1.5">
            {SUPPORT_WALLETS.map((w) => (
              <div key={w.asset} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-2.5 py-1.5">
                <span className="w-32 shrink-0 text-xs font-medium text-[var(--text-dim)]">{w.asset}</span>
                <span className="min-w-0 flex-1 truncate font-mono text-xs text-[var(--text)]" title={w.address}>{w.address}</span>
                <button
                  onClick={() => { navigator.clipboard.writeText(w.address); toast('ok', t('settings.supportCopied', { a: w.asset })) }}
                  className="shrink-0 rounded-md p-1 text-[var(--text-faint)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--text)]"
                  title={t('common.copy')}>
                  <Copy size={13} />
                </button>
              </div>
            ))}
          </div>
        </Card>

        <Card title={t('settings.about')}>
          <p className="text-sm text-[var(--text-dim)]">🗡️ SwissKnife for MS Graph</p>
          <p className="mt-1 text-sm text-[var(--text-faint)]">{t('settings.version', { v: version || '—' })}</p>
          <div className="mt-3 flex items-center gap-3">
            <Button variant="subtle" onClick={checkUpdates} disabled={checking}>
              {checking ? <Spinner /> : <RefreshCw size={15} />} {checking ? t('settings.checking') : t('settings.checkUpdates')}
            </Button>
            {update && !update.updateAvailable && (
              // A dev build can't be meaningfully compared to release tags —
              // say what the latest release is instead of claiming "latest".
              update.currentVersion.includes('dev') && update.latestVersion
                ? <span className="text-sm text-[var(--text-dim)]">{t('settings.devBuild', { v: update.latestVersion })}</span>
                : <span className="text-sm text-[var(--ok)]">{t('settings.upToDate')}</span>
            )}
            {update && update.updateAvailable && (
              <div className="flex flex-wrap items-center gap-2">
                <Badge kind="warn">{t('settings.updateAvailable', { v: update.latestVersion })}</Badge>
                {update.assetUrl && (
                  <Button variant="primary" disabled={updating} onClick={updateNow}>
                    {updating ? <Spinner /> : <Download size={15} />} {updating ? (updateProgress || t('common.working')) : t('settings.updateNow')}
                  </Button>
                )}
                <Button variant="subtle" disabled={updating} onClick={() => api.update.openReleases(update.url)}>
                  {t('settings.download')}
                </Button>
              </div>
            )}
          </div>
        </Card>
      </div>
    </Page>
  )
}
