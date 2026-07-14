import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { UserPlus, UserMinus, Play, CheckCircle2, XCircle } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Field, Input, Spinner } from '../components/ui'
import { MultiSelect, type Option } from '../components/MultiSelect'
import { EntityPicker } from '../components/EntityPicker'
import { UpnInput } from '../components/UpnInput'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import { skuFriendly } from '../lib/skuNames'
import type { services } from '../../wailsjs/go/models'

export function PlaybooksPage() {
  const { t } = useTranslation()
  const { readOnly, connected, toast } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const [tab, setTab] = useState<'onboard' | 'offboard'>('onboard')
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState<services.PlaybookResult | null>(null)

  const [on, setOn] = useState({ displayName: '', upn: '', mailNickname: '', password: '', usageLocation: '' })
  const [skuIds, setSkuIds] = useState<string[]>([])
  const [groupIds, setGroupIds] = useState<string[]>([])
  const [teamIds, setTeamIds] = useState<string[]>([])
  const [off, setOff] = useState({ upn: '', block: true, revokeSessions: true, removeAllLicenses: true, backupToUser: '', backupFolder: '', delete: false })

  const [skuOpts, setSkuOpts] = useState<Option[]>([])
  const [groupOpts, setGroupOpts] = useState<Option[]>([])
  const [teamOpts, setTeamOpts] = useState<Option[]>([])
  const [loadingOpts, setLoadingOpts] = useState(false)

  // Channel picker (nested: pick a team, then its channels).
  const [chTeam, setChTeam] = useState('')
  const [chOpts, setChOpts] = useState<Option[]>([])
  const [chSelected, setChSelected] = useState<string[]>([])
  const [chLoading, setChLoading] = useState(false)

  useEffect(() => {
    if (!connected) return
    setLoadingOpts(true)
    Promise.allSettled([api.licensing.skus(), api.groups.list('', 0), api.teams.all()]).then(([s, g, tm]) => {
      if (s.status === 'fulfilled') setSkuOpts((s.value as GraphObject[]).map((x) => ({ value: x.skuId, label: skuFriendly(x.skuPartNumber), sub: `${x.consumedUnits ?? ''}/${x.prepaidUnits?.enabled ?? ''}` })))
      if (g.status === 'fulfilled') setGroupOpts((g.value as GraphObject[]).map((x) => ({ value: x.id, label: x.displayName || x.id, sub: x.mail })))
      if (tm.status === 'fulfilled') setTeamOpts((tm.value as GraphObject[]).map((x) => ({ value: x.id, label: x.displayName || x.id })))
      setLoadingOpts(false)
    })
  }, [connected])

  // Load channels when a team is chosen for the channel picker.
  useEffect(() => {
    setChOpts([]); setChSelected([])
    if (!chTeam) return
    setChLoading(true)
    api.teams.channels(chTeam)
      .then((chs) => setChOpts((chs as GraphObject[]).map((c) => ({ value: c.id, label: c.displayName || c.id, sub: c.membershipType }))))
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setChLoading(false))
  }, [chTeam])

  const runOnboard = async () => {
    setBusy(true); setResult(null)
    try {
      const channelRefs = chSelected.map((cid) => ({ teamId: chTeam, channelId: cid }))
      setResult(await api.playbooks.onboard({
        displayName: on.displayName, upn: on.upn, mailNickname: on.mailNickname, password: on.password,
        usageLocation: on.usageLocation, skuIds, groupIds, teamIds, channelRefs,
      }))
      toast('ok', t('playbooks.onboard'))
    } catch (e) { toast('err', errMessage(e)) } finally { setBusy(false) }
  }

  const runOffboard = (confirm: string) => async () => {
    setBusy(true); setResult(null)
    try {
      setResult(await api.playbooks.offboard({ ...off, confirm }))
      toast('ok', t('playbooks.offboard'))
    } catch (e) { toast('err', errMessage(e)) } finally { setBusy(false) }
  }

  return (
    <Page title={t('playbooks.title')} subtitle={t('playbooks.subtitle')}>
      {confirmElement}
      <div className="mb-4 inline-flex rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
        {(['onboard', 'offboard'] as const).map((x) => (
          <button key={x} onClick={() => { setTab(x); setResult(null) }}
            className={`flex items-center gap-2 rounded-md px-4 py-1.5 text-sm font-medium ${tab === x ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)]'}`}>
            {x === 'onboard' ? <UserPlus size={15} /> : <UserMinus size={15} />}
            {t(`playbooks.${x}`)}
          </button>
        ))}
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-[minmax(380px,460px)_1fr]">
        {tab === 'onboard' ? (
          <Card title={t('playbooks.onboard')}>
            <div className="flex flex-col gap-2.5">
              <Input placeholder={t('playbooks.displayName')} value={on.displayName} onChange={(e) => setOn({ ...on, displayName: e.target.value })} />
              <Input placeholder="user@contoso.com" value={on.upn} onChange={(e) => setOn({ ...on, upn: e.target.value })} />
              <Input placeholder={t('playbooks.mailNickname')} value={on.mailNickname} onChange={(e) => setOn({ ...on, mailNickname: e.target.value })} />
              <Input type="password" placeholder={t('playbooks.password')} value={on.password} onChange={(e) => setOn({ ...on, password: e.target.value })} />
              <Input placeholder={t('playbooks.usageLocation')} value={on.usageLocation} onChange={(e) => setOn({ ...on, usageLocation: e.target.value })} />
              <Field label={t('playbooks.licenses')}><MultiSelect options={skuOpts} selected={skuIds} onChange={setSkuIds} loading={loadingOpts} placeholder={t('playbooks.pickLicenses')} /></Field>
              <Field label={t('playbooks.groups')}><MultiSelect options={groupOpts} selected={groupIds} onChange={setGroupIds} loading={loadingOpts} placeholder={t('playbooks.pickGroups')} /></Field>
              <Field label={t('playbooks.teams')}><MultiSelect options={teamOpts} selected={teamIds} onChange={setTeamIds} loading={loadingOpts} placeholder={t('playbooks.pickTeams')} /></Field>
              <Field label={t('playbooks.channels')}>
                <div className="flex flex-col gap-2">
                  <EntityPicker value={chTeam} onChange={setChTeam} load={async () => teamOpts} placeholder={t('playbooks.pickTeamForChannels')} />
                  {chTeam && <MultiSelect options={chOpts} selected={chSelected} onChange={setChSelected} loading={chLoading} placeholder={t('playbooks.pickChannels')} />}
                </div>
              </Field>
              <Button variant="primary" disabled={readOnly || busy || !on.upn || !on.displayName || !on.mailNickname || !on.password} onClick={runOnboard}>
                {busy ? <Spinner /> : <Play size={15} />} {t('playbooks.run')}
              </Button>
            </div>
          </Card>
        ) : (
          <Card title={t('playbooks.offboard')}>
            <div className="flex flex-col gap-2">
              <Field label={t('playbooks.upn')}><UpnInput value={off.upn} onChange={(v) => setOff({ ...off, upn: v })} /></Field>
              {([['block', 'block'], ['revokeSessions', 'revoke'], ['removeAllLicenses', 'removeLicenses'], ['delete', 'deleteUser']] as const).map(([k, label]) => (
                <label key={k} className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                  <input type="checkbox" checked={(off as any)[k]} onChange={(e) => setOff({ ...off, [k]: e.target.checked })} />
                  {t(`playbooks.${label}`)}
                </label>
              ))}
              <Field label={t('playbooks.backupTo')}><UpnInput value={off.backupToUser} onChange={(v) => setOff({ ...off, backupToUser: v })} placeholder="backup@contoso.com" /></Field>
              <Input placeholder={t('playbooks.backupFolder')} value={off.backupFolder} onChange={(e) => setOff({ ...off, backupFolder: e.target.value })} />
              <Button variant="danger" disabled={readOnly || busy || !off.upn} onClick={() => askConfirm(off.upn, (c) => runOffboard(c)())}>
                {busy ? <Spinner /> : <Play size={15} />} {t('playbooks.run')}
              </Button>
            </div>
          </Card>
        )}

        <Card title={t('playbooks.steps')}>
          {!result && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          <div className="flex flex-col gap-2">
            {result?.steps?.map((s, i) => (
              <div key={i} className="flex items-start gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm">
                {s.ok ? <CheckCircle2 size={16} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={16} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                <div className="min-w-0">
                  <div className="font-medium text-[var(--text)]">{s.name}{s.detail ? <span className="ml-2 text-xs text-[var(--text-faint)]">{s.detail}</span> : null}</div>
                  {s.error && <div className="text-xs text-[var(--danger)]">{s.error}</div>}
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Page>
  )
}
