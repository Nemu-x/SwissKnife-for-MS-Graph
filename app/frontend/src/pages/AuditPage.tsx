import { useTranslation } from 'react-i18next'
import { LogIn, FileClock } from 'lucide-react'
import { ActionPage, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { useAsync } from '../lib/useAsync'
import { api, type GraphObject } from '../lib/api'

export function AuditPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[]>()

  const actions: Action[] = [
    { id: 'signins', label: 'Sign-in logs', icon: <LogIn size={15} />, variant: 'primary', onClick: () => res.run(() => api.audit.signIns(50)) },
    { id: 'directory', label: 'Directory audits', icon: <FileClock size={15} />, onClick: () => res.run(() => api.audit.directory(50)) },
  ]

  return <ActionPage title={t('nav.audit')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} />} />
}
