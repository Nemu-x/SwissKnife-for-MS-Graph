import { useTranslation } from 'react-i18next'
import { Heart, Star, Copy, ChevronDown } from 'lucide-react'
import { Card } from './ui'
import { useStore } from '../lib/store'
import { BrowserOpenURL, ClipboardSetText } from '../../wailsjs/runtime/runtime'
import { GITHUB_URL, SUPPORT_WALLETS, TRIBUTE_URL } from '../lib/support'

// Supporting the project: Tribute first (a card or Telegram, once or monthly),
// a GitHub star as the free option, crypto folded away for those who want it.
export function SupportCard() {
  const { t } = useTranslation()
  const { toast } = useStore()
  const copy = (asset: string, address: string) =>
    ClipboardSetText(address)
      .then((ok) => toast(ok ? 'ok' : 'err', ok ? t('settings.supportCopied', { a: asset }) : t('common.copyFailed')))
      .catch(() => toast('err', t('common.copyFailed')))

  return (
    <Card title={t('support.title')}>
      <div className="flex flex-col gap-3">
        <p className="text-sm leading-relaxed text-[var(--text-dim)]">{t('support.pitch')}</p>
        <button onClick={() => BrowserOpenURL(TRIBUTE_URL)}
          className="group flex items-center gap-3 rounded-xl border border-[var(--accent)]/40 bg-[var(--accent)]/10 px-4 py-3 text-left transition-colors hover:bg-[var(--accent)]/20">
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-[var(--accent)] text-[var(--accent-fg)]">
            <Heart size={18} className="transition-transform group-hover:scale-110" />
          </span>
          <span className="min-w-0 flex-1">
            <span className="block text-sm font-semibold">{t('support.tribute')}</span>
            <span className="block text-xs text-[var(--text-faint)]">{t('support.tributeHint')}</span>
          </span>
        </button>
        <button onClick={() => BrowserOpenURL(GITHUB_URL)}
          className="flex items-center gap-3 rounded-xl border border-[var(--border)] px-4 py-2.5 text-left hover:bg-[var(--bg-elev-2)]">
          <Star size={16} className="shrink-0 text-[var(--warn)]" />
          <span className="min-w-0 flex-1">
            <span className="block text-sm">{t('support.star')}</span>
            <span className="block text-xs text-[var(--text-faint)]">{t('support.starHint')}</span>
          </span>
        </button>
        <details className="group rounded-xl border border-[var(--border)] px-4 py-2.5">
          <summary className="flex cursor-pointer list-none items-center justify-between text-sm text-[var(--text-dim)]">
            {t('support.crypto')}
            <ChevronDown size={14} className="transition-transform group-open:rotate-180" />
          </summary>
          <div className="mt-2 flex flex-col gap-1.5">
            {SUPPORT_WALLETS.map((w) => (
              <div key={w.asset} className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-2.5 py-1.5">
                <span className="w-28 shrink-0 text-xs font-medium text-[var(--text-dim)]">{w.asset}</span>
                <span className="min-w-0 flex-1 truncate font-mono text-xs" title={w.address}>{w.address}</span>
                <button onClick={() => copy(w.asset, w.address)} aria-label={`${t('common.copy')} ${w.asset}`}
                  className="shrink-0 rounded-md p-1 text-[var(--text-faint)] hover:bg-[var(--bg-elev-2)] hover:text-[var(--text)]">
                  <Copy size={13} />
                </button>
              </div>
            ))}
            <p className="text-[11px] text-[var(--text-faint)]">{t('support.cryptoWarn')}</p>
          </div>
        </details>
      </div>
    </Card>
  )
}
