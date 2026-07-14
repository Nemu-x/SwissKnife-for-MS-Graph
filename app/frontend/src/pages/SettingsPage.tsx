import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Moon, Sun, Lock } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Select, Button } from '../components/ui'
import { useStore } from '../lib/store'
import { setLanguage } from '../i18n'
import { api } from '../lib/api'
import { Version } from '../../wailsjs/go/main/App'

export function SettingsPage() {
  const { t, i18n } = useTranslation()
  const { theme, toggleTheme, safeMode, setSafeMode, readOnly, setStatus, connected } = useStore()
  const [version, setVersion] = useState('')

  useEffect(() => { Version().then(setVersion).catch(() => {}) }, [])

  const toggleReadOnly = async () => setStatus(await api.connect.setReadOnly(!readOnly))

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

        <Card title={t('settings.about')}>
          <p className="text-sm text-[var(--text-dim)]">🗡️ SwissKnife for MS Graph</p>
          <p className="mt-1 text-sm text-[var(--text-faint)]">{t('settings.version', { v: version || '—' })}</p>
        </Card>
      </div>
    </Page>
  )
}
