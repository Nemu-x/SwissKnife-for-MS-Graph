import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { RefreshCw, Download, ExternalLink } from 'lucide-react'
import { Card, Button, Badge, Spinner } from './ui'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import { BrowserOpenURL } from '../../wailsjs/runtime/runtime'
import type { services } from '../../wailsjs/go/models'

// PowerShell backends (ADR-008): what is installed, and one-click installs of
// the modules the Exchange and Teams actions need. Detection spawns pwsh, so
// it runs when the card is shown, not at app start.
export function PowerShellCard() {
  const { t } = useTranslation()
  const { toast, jobs, patchJob } = useStore()
  const [st, setSt] = useState<services.PowerShellStatus | null>(null)
  // An install takes minutes: it lives in the store's job slot so leaving
  // Settings and coming back neither loses it nor allows a second one.
  const job = jobs.psInstall
  const installing = job?.running ? job.progress : null
  const checking = !!jobs.psCheck?.running
  const busy = installing || (checking ? 'refresh' : null)

  useEffect(() => {
    let alive = true
    api.actions.powerShellStatus().then((s) => alive && setSt(s)).catch(() => {})
    return () => { alive = false }
  }, [job?.running, jobs.psCheck?.running])

  const refresh = async () => {
    if (jobs.psCheck?.running) return
    patchJob('psCheck', { running: true, startedAt: Date.now() })
    try { setSt(await api.actions.refreshPowerShell()) } finally { patchJob('psCheck', { running: false }) }
  }

  const install = async (name: string) => {
    if (jobs.psInstall?.running) return
    patchJob('psInstall', { running: true, progress: name, error: null, startedAt: Date.now() })
    try {
      setSt(await api.actions.installModule(name))
      toast('ok', t('powershell.installed', { name }))
    } catch (e) {
      const m = errMessage(e)
      patchJob('psInstall', { error: m })
      toast('err', m)
    } finally { patchJob('psInstall', { running: false, progress: '' }) }
  }

  return (
    <Card title={t('powershell.title')}>
      <p className="mb-3 text-xs leading-relaxed text-[var(--text-faint)]">{t('powershell.why')}</p>
      {!st && <div className="flex items-center gap-2 text-sm text-[var(--text-dim)]"><Spinner /> {t('powershell.detecting')}</div>}
      {st && !st.installed && (
        <div className="flex flex-col gap-2 text-sm">
          <p className="text-[var(--warn)]">{st.exe ? t('powershell.tooOld', { version: st.version }) : t('powershell.missing')}</p>
          <button onClick={() => BrowserOpenURL('https://aka.ms/powershell-release?tag=stable')}
            className="flex items-center gap-1.5 text-left text-[var(--accent)] hover:underline">
            <ExternalLink size={14} /> {t('powershell.getIt')}
          </button>
        </div>
      )}
      {st?.installed && (
        <div className="flex flex-col gap-2">
          <div className="text-sm text-[var(--text-dim)]">PowerShell {st.version}</div>
          {st.modules.map((m) => (
            <div key={m.name} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2">
              <span className="min-w-0 flex-1 truncate font-mono text-xs">{m.name}</span>
              {m.usable
                ? <Badge kind="ok">{m.version}</Badge>
                : (
                  <Button variant="subtle" className="!px-2 !py-1" disabled={!!busy} onClick={() => install(m.name)}>
                    {busy === m.name ? <Spinner /> : <Download size={14} />} {m.version ? t('powershell.update') : t('powershell.install')}
                  </Button>
                )}
            </div>
          ))}
          {busy && busy !== 'refresh' && <p className="text-xs text-[var(--text-faint)]">{t('powershell.installing')}</p>}
        </div>
      )}
      <Button variant="ghost" className="mt-3 !px-2 !py-1" disabled={!!busy} onClick={refresh}>
        {busy === 'refresh' ? <Spinner /> : <RefreshCw size={14} />} {t('powershell.recheck')}
      </Button>
    </Card>
  )
}
