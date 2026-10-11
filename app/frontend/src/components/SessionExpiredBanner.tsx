import { useTranslation } from 'react-i18next'
import { KeyRound, X } from 'lucide-react'
import { Button } from './ui'
import { useStore } from '../lib/store'

// Shown app-wide when any call reports that the delegated sign-in expired:
// the operator signs in again on the Connect page instead of every page
// failing on its own.
export function SessionExpiredBanner({ onSignIn }: { onSignIn: () => void }) {
  const { t } = useTranslation()
  const { sessionExpired, dismissSessionExpired } = useStore()
  if (!sessionExpired) return null
  return (
    <div role="alert" data-testid="session-expired"
      className="fixed left-1/2 top-3 z-50 flex max-w-xl -translate-x-1/2 items-center gap-3 rounded-lg border border-[var(--warn)]/50 bg-[var(--bg-elev)] px-4 py-2 shadow-lg">
      <KeyRound size={16} className="shrink-0 text-[var(--warn)]" />
      <span className="text-sm text-[var(--text)]">{t('session.expired')}</span>
      <Button variant="primary" className="!px-2 !py-1" onClick={onSignIn}>{t('session.signInAgain')}</Button>
      <button onClick={dismissSessionExpired} aria-label={t('common.close')} className="text-[var(--text-faint)] hover:text-[var(--text)]">
        <X size={14} />
      </button>
    </div>
  )
}
