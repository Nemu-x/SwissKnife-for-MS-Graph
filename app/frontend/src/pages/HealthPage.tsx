import { useTranslation } from 'react-i18next'
import { Activity, AlertTriangle, Megaphone } from 'lucide-react'
import { TaskPage, type TaskAction } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { useAsync } from '../lib/useAsync'
import { api, type GraphObject } from '../lib/api'

export function HealthPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[]>()

  const actions: TaskAction[] = [
    {
      id: 'overview', label: t('health.tileOverview'), hint: t('health.hintOverview'), icon: <Activity size={16} />, variant: 'primary',
      onClick: () => res.run(() => api.health.overview()),
    },
    {
      id: 'issues', label: t('health.tileIssues'), hint: t('health.hintIssues'), icon: <AlertTriangle size={16} />, variant: 'primary',
      onClick: () => res.run(() => api.health.issues()),
    },
    {
      id: 'messages', label: t('health.tileMessages'), hint: t('health.hintMessages'), icon: <Megaphone size={16} />,
      onClick: () => res.run(() => api.health.messages()),
    },
  ]

  return (
    <TaskPage
      pageId="health"
      title={t('health.title')}
      subtitle={t('health.subtitle')}
      actions={actions}
      busy={res.loading}
      onClearResult={res.reset}
      hasResult={!!res.data || res.loading || !!res.error}
      result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
    />
  )
}
