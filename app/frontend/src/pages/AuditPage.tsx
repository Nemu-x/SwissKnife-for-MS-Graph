import { useTranslation } from 'react-i18next'
import { LogIn, FileClock } from 'lucide-react'
import { Page } from '../components/Layout'
import { TwoPane } from '../components/TwoPane'
import { ResultView } from '../components/ResultView'
import { Button, Card } from '../components/ui'
import { useAsync } from '../lib/useAsync'
import { type GraphObject } from '../lib/api'
import { api } from '../lib/api'

export function AuditPage() {
  const { t } = useTranslation()
  const res = useAsync<GraphObject[]>()

  return (
    <Page title={t('nav.audit')}>
      <TwoPane
        controls={
          <Card title={t('nav.audit')}>
            <div className="flex flex-col gap-2">
              <Button variant="primary" onClick={() => res.run(() => api.audit.signIns(50))}>
                <LogIn size={15} /> Sign-in logs
              </Button>
              <Button variant="subtle" onClick={() => res.run(() => api.audit.directory(50))}>
                <FileClock size={15} /> Directory audits
              </Button>
            </div>
          </Card>
        }
        result={<ResultView data={res.data} loading={res.loading} error={res.error} />}
      />
    </Page>
  )
}
