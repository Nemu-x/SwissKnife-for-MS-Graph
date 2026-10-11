import { useTranslation } from 'react-i18next'
import { Badge } from './ui'

const DAY = 24 * 60 * 60 * 1000
// Matches auth.CertExpiringWindow on the backend.
const WARN_DAYS = 30

// Warns on a certificate profile whose certificate expires within 30 days or
// already has. certNotAfter comes from the profile list (read from the PFX).
export function CertExpiryBadge({ notAfter }: { notAfter?: string | null }) {
  const { t } = useTranslation()
  if (!notAfter) return null
  const end = new Date(notAfter)
  if (isNaN(end.getTime())) return null
  const days = Math.floor((end.getTime() - Date.now()) / DAY)
  if (days > WARN_DAYS) return null
  const date = end.toLocaleDateString()
  if (days < 0) {
    return <span title={t('connect.certExpiredHint', { date })}><Badge kind="danger">{t('connect.certExpired')}</Badge></span>
  }
  return <span title={t('connect.certExpiringHint', { date })}><Badge kind="warn">{t('connect.certExpiresIn', { n: days })}</Badge></span>
}
