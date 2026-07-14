import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Moon, Sun, Lock, RefreshCw, Download } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Select, Button, Badge, Spinner } from '../components/ui'
import { useStore } from '../lib/store'
import { ACCENT_PRESETS, isHex } from '../lib/color'
import { setLanguage } from '../i18n'
import { api, errMessage } from '../lib/api'
import { Version } from '../../wailsjs/go/main/App'
import type { services } from '../../wailsjs/go/models'

export function SettingsPage() {
  const { t, i18n } = useTranslation()
  const { theme, toggleTheme, accent, setAccent, safeMode, setSafeMode, hideUnavailable, setHideUnavailable, checkAccess, readOnly, setStatus, connected, toast } = useStore()
  const [checkingAccess, setCheckingAccess] = useState(false)
  const [version, setVersion] = useState('')
  const [update, setUpdate] = useState<services.UpdateInfo | null>(null)
  const [checking, setChecking] = useState(false)

  useEffect(() => { Version().then(setVersion).catch(() => {}) }, [])

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

        <Card title={t('settings.about')}>
          <p className="text-sm text-[var(--text-dim)]">🗡️ SwissKnife for MS Graph</p>
          <p className="mt-1 text-sm text-[var(--text-faint)]">{t('settings.version', { v: version || '—' })}</p>
          <div className="mt-3 flex items-center gap-3">
            <Button variant="subtle" onClick={checkUpdates} disabled={checking}>
              {checking ? <Spinner /> : <RefreshCw size={15} />} {checking ? t('settings.checking') : t('settings.checkUpdates')}
            </Button>
            {update && !update.updateAvailable && (
              <span className="text-sm text-[var(--ok)]">{t('settings.upToDate')}</span>
            )}
            {update && update.updateAvailable && (
              <div className="flex items-center gap-2">
                <Badge kind="warn">{t('settings.updateAvailable', { v: update.latestVersion })}</Badge>
                <Button variant="primary" onClick={() => api.update.openReleases(update.url)}>
                  <Download size={15} /> {t('settings.download')}
                </Button>
              </div>
            )}
          </div>
        </Card>
      </div>
    </Page>
  )
}
