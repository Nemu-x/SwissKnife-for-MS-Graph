import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { UserPlus, UserMinus, Play, CheckCircle2, XCircle, Loader2, X } from 'lucide-react'
import { Page } from '../components/Layout'
import { Card, Button, Field, Input, Spinner, ErrorNote, Badge } from '../components/ui'
import { MultiSelect, type Option } from '../components/MultiSelect'
import { EntityPicker } from '../components/EntityPicker'
import { UpnInput } from '../components/UpnInput'
import { useConfirm } from '../lib/useConfirm'
import { useStore, type PlaybookLiveStep } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import { skuFriendly } from '../lib/skuNames'
import type { services } from '../../wailsjs/go/models'

export function PlaybooksPage() {
  const { t } = useTranslation()
  const { readOnly, connected, toast, jobs, startPlaybook, cancelPlaybook, clearJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const [tab, setTab] = useState<'onboard' | 'offboard'>('onboard')

  // The run lives in the store, so steps and the final report survive
  // navigating away and back while a playbook is executing.
  const job = jobs.playbook
  const busy = !!job?.running
  const result = (job?.result as services.PlaybookResult | null) ?? null
  const steps: PlaybookLiveStep[] = result?.steps ?? job?.steps ?? []

  const [on, setOn] = useState({ displayName: '', upn: '', mailNickname: '', password: '', usageLocation: '' })
  const [skuIds, setSkuIds] = useState<string[]>([])
  const [groupIds, setGroupIds] = useState<string[]>([])
  const [teamIds, setTeamIds] = useState<string[]>([])
  const [off, setOff] = useState({
    upn: '', block: true, revokeSessions: true,
    oof: false, oofMessage: '', forwardTo: '', hideFromGal: false, calendarTo: '', removeFromGroups: false,
    removeAllLicenses: true, backupToUser: '', backupFolder: '', delete: false,
  })

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

  const runOnboard = () => {
    const channelRefs = chSelected.map((cid) => ({ teamId: chTeam, channelId: cid }))
    startPlaybook('onboard', on.upn, () => api.playbooks.onboard({
      displayName: on.displayName, upn: on.upn, mailNickname: on.mailNickname, password: on.password,
      usageLocation: on.usageLocation, skuIds, groupIds, teamIds, channelRefs,
    }))
  }

  const runOffboard = (confirm: string) => () => {
    startPlaybook('offboard', off.upn, () => api.playbooks.offboard({ ...off, confirm }))
  }

  return (
    <Page title={t('playbooks.title')} subtitle={t('playbooks.subtitle')}>
      {confirmElement}
      <div className="mb-4 inline-flex rounded-lg border border-[var(--border)] bg-[var(--bg-elev)] p-1">
        {(['onboard', 'offboard'] as const).map((x) => (
          <button key={x} onClick={() => { setTab(x); if (!busy) clearJob('playbook') }}
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
              {([['block', 'block'], ['revokeSessions', 'revoke'], ['oof', 'oof'], ['hideFromGal', 'hideFromGal'], ['removeFromGroups', 'removeFromGroups'], ['removeAllLicenses', 'removeLicenses'], ['delete', 'deleteUser']] as const).map(([k, label]) => (
                <div key={k}>
                  <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
                    <input type="checkbox" checked={(off as any)[k]} onChange={(e) => setOff({ ...off, [k]: e.target.checked })} />
                    {t(`playbooks.${label}`)}
                  </label>
                  {/* Mailbox-safety warnings: license removal and account deletion
                      destroy mail unless the mailbox was converted to shared first. */}
                  {k === 'removeAllLicenses' && off.removeAllLicenses && (
                    <p className="ml-6 mt-0.5 text-xs text-[var(--warn)]">{t('playbooks.removeLicensesWarn')}</p>
                  )}
                  {k === 'delete' && off.delete && (
                    <p className="ml-6 mt-0.5 text-xs text-[var(--danger)]">{t('playbooks.deleteUserWarn')}</p>
                  )}
                </div>
              ))}
              {off.oof && (
                <Input placeholder={t('playbooks.oofMessage')} value={off.oofMessage} onChange={(e) => setOff({ ...off, oofMessage: e.target.value })} />
              )}
              <Field label={t('playbooks.forwardTo')}><UpnInput value={off.forwardTo} onChange={(v) => setOff({ ...off, forwardTo: v })} placeholder="manager@contoso.com" /></Field>
              <Field label={t('playbooks.calendarTo')}><UpnInput value={off.calendarTo} onChange={(v) => setOff({ ...off, calendarTo: v })} placeholder="manager@contoso.com" /></Field>
              <Field label={t('playbooks.backupTo')}><UpnInput value={off.backupToUser} onChange={(v) => setOff({ ...off, backupToUser: v })} placeholder="backup@contoso.com" /></Field>
              <Input placeholder={t('playbooks.backupFolder')} value={off.backupFolder} onChange={(e) => setOff({ ...off, backupFolder: e.target.value })} />
              <Button variant="danger" disabled={readOnly || busy || !off.upn} onClick={() => askConfirm(off.upn, (c) => runOffboard(c)())}>
                {busy ? <Spinner /> : <Play size={15} />} {t('playbooks.run')}
              </Button>
            </div>
          </Card>
        )}

        <Card title={t('playbooks.steps')}>
          {busy && (
            <div className="mb-3 flex items-start justify-between gap-3">
              <p className="flex items-center gap-2 text-sm text-[var(--accent2)]">
                <Loader2 size={14} className="shrink-0 animate-spin" /> {t('playbooks.runningNote')}
              </p>
              <button onClick={cancelPlaybook} disabled={job?.canceled}
                className="flex shrink-0 items-center gap-1 rounded-md border border-[var(--danger)]/40 px-2 py-0.5 text-xs text-[var(--danger)] hover:bg-[var(--danger)]/10 disabled:opacity-50">
                <X size={12} /> {job?.canceled ? t('common.canceling') : t('common.cancel')}
              </button>
            </div>
          )}
          {job?.error && <ErrorNote>{job.error}</ErrorNote>}
          {result?.canceled && <div className="mb-2"><Badge kind="warn">{t('common.canceled')}</Badge></div>}
          {steps.length === 0 && !busy && !job?.error && <p className="text-sm text-[var(--text-faint)]">{t('common.empty')}</p>}
          <div className="flex flex-col gap-2">
            {steps.map((s, i) => (
              <div key={i} className="flex items-start gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm">
                {s.running
                  ? <Loader2 size={16} className="mt-0.5 shrink-0 animate-spin text-[var(--accent2)]" />
                  : s.ok
                    ? <CheckCircle2 size={16} className="mt-0.5 shrink-0 text-[var(--ok)]" />
                    : <XCircle size={16} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                <div className="min-w-0">
                  <div className="font-medium text-[var(--text)]">{s.name}{s.detail ? <span className="ml-2 text-xs text-[var(--text-faint)]">{s.detail}</span> : null}</div>
                  {/* Live percentage for the step in flight (e.g. the OneDrive backup). */}
                  {s.running && job?.progress && <div className="text-xs text-[var(--accent2)]">{job.progress}</div>}
                  {s.error && <div className="text-xs text-[var(--danger)]">{s.errorCode ? `${s.errorCode}: ` : ''}{s.error}</div>}
                  {s.hint && <div className="text-xs text-[var(--warn)]">{t('playbooks.permissionHint', { p: s.hint })}</div>}
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Page>
  )
}
