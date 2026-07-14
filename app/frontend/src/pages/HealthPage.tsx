import { useTranslation } from 'react-i18next'
import { Activity, AlertTriangle, Megaphone } from 'lucide-react'
import { ActionPage, type Action } from '../components/ActionPage'
import { ResultView } from '../components/ResultView'
import { useAsync } from '../lib/useAsync'
import { api, type GraphObject } from '../lib/api'

export function HealthPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[]>()

  const actions: Action[] = [
    { id: 'overview', label: t('health.overview'), icon: <Activity size={15} />, variant: 'primary', onClick: () => res.run(() => api.health.overview()) },
    { id: 'issues', label: t('health.issues'), icon: <AlertTriangle size={15} />, onClick: () => res.run(() => api.health.issues()) },
    { id: 'messages', label: t('health.messages'), icon: <Megaphone size={15} />, onClick: () => res.run(() => api.health.messages()) },
  ]

  return <ActionPage title={t('health.title')} actions={actions} result={<ResultView data={res.data} loading={res.loading} error={res.error} />} />
}
