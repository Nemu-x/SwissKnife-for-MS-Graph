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
import { humanBytes } from '../lib/format'
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

  // Backend emits stable keys + params; the English text in name/detail is the
  // fallback for keys this build doesn't know yet.
  const stepName = (s: PlaybookLiveStep) => (s.nameKey ? t(s.nameKey, { defaultValue: s.name }) : s.name)
  const stepDetail = (s: PlaybookLiveStep) => {
    if (!s.detailKey) return s.detail
    const p = s.params || {}
    let out = t(s.detailKey, { ...p, size: humanBytes(p.bytes || 0), defaultValue: s.detail || '' })
    if (s.detailKey === 'stepDetails.backup') {
      if (p.skipped > 0) out += t('stepDetails.skippedSuffix', { n: p.skipped })
      if (p.canceled) out += t('stepDetails.canceledSuffix')
    }
    return out
  }

  const [on, setOn] = useState({ displayName: '', upn: '', mailNickname: '', password: '', usageLocation: '' })
  const [skuIds, setSkuIds] = useState<string[]>([])
  const [groupIds, setGroupIds] = useState<string[]>([])
  const [teamIds, setTeamIds] = useState<string[]>([])
  // Frequently reused targets are remembered from the last successful run.
  const remembered = (k: string) => localStorage.getItem('defaults.offboard.' + k) || ''
  const [off, setOff] = useState({
    upn: '', block: true, revokeSessions: true,
    oof: false, oofMessage: '', forwardTo: remembered('forwardTo'), hideFromGal: false,
    calendarTo: remembered('calendarTo'), removeFromGroups: false,
    removeAllLicenses: true, backupToUser: remembered('backupToUser'), backupFolder: remembered('backupFolder'),
    backupChats: false, delete: false,
  })

  // Named option presets; two built-ins encode the blessed two-phase flow
  // (convert-to-shared happens manually between the phases).
  type OffOptions = Omit<typeof off, 'upn'>
  const BUILTIN_PRESETS: Record<string, Partial<OffOptions>> = {
    __phase1: {
      block: true, revokeSessions: true, hideFromGal: true, removeFromGroups: true,
      removeAllLicenses: false, delete: false,
    },
    __phase2: {
      block: false, revokeSessions: false, oof: false, hideFromGal: false, removeFromGroups: false,
      removeAllLicenses: true, delete: false, forwardTo: '', calendarTo: '', backupToUser: '', backupFolder: '',
    },
  }
  const loadPresets = (): Record<string, Partial<OffOptions>> => {
    // Persisted data is untrusted: accept only an object of objects.
    try {
      const raw = JSON.parse(localStorage.getItem('playbook.presets') || '{}')
      if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {}
      const out: Record<string, Partial<OffOptions>> = {}
      for (const [k, v] of Object.entries(raw)) {
        if (v && typeof v === 'object' && !Array.isArray(v)) out[k] = v as Partial<OffOptions>
      }
      return out
    } catch { return {} }
  }
  const [presets, setPresets] = useState(loadPresets)
  const [presetSel, setPresetSel] = useState('')
  const [presetName, setPresetName] = useState('')
  // Own-property check: a custom preset named "toString" must not resolve to
  // the inherited Object.prototype member.
  const isBuiltinPreset = (name: string) => Object.prototype.hasOwnProperty.call(BUILTIN_PRESETS, name)
  const applyPreset = (name: string) => {
    setPresetSel(name)
    const p = isBuiltinPreset(name) ? BUILTIN_PRESETS[name] : presets[name]
    if (p) setOff((o) => ({ ...o, ...p }))
  }
  const savePreset = () => {
    const name = presetName.trim()
    if (!name || name.startsWith('__')) return
    const { upn: _upn, ...opts } = off
    const next = { ...presets, [name]: opts }
    setPresets(next)
    localStorage.setItem('playbook.presets', JSON.stringify(next))
    setPresetSel(name); setPresetName('')
  }
  const deletePreset = () => {
    if (!presetSel || isBuiltinPreset(presetSel)) return
    const next = { ...presets }
    delete next[presetSel]
    setPresets(next)
    localStorage.setItem('playbook.presets', JSON.stringify(next))
    setPresetSel('')
  }

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
    startPlaybook('offboard', off.upn, () => api.playbooks.offboard({ ...off, confirm })).then((r) => {
      if (!r || r.canceled) return // failed or canceled run: keep the previous defaults
      for (const k of ['forwardTo', 'calendarTo', 'backupToUser', 'backupFolder'] as const) {
        // A run that skipped a field must not erase the remembered value.
        if (off[k]) localStorage.setItem('defaults.offboard.' + k, off[k])
      }
    })
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
              <Field label={t('playbooks.presets')}>
                <div className="flex items-center gap-2">
                  <select value={presetSel} onChange={(e) => applyPreset(e.target.value)}
                    className="min-w-0 flex-1 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-2 py-1.5 text-sm outline-none focus:border-[var(--accent)]">
                    <option value="">{t('playbooks.presetPick')}</option>
                    <option value="__phase1">{t('playbooks.presetPhase1')}</option>
                    <option value="__phase2">{t('playbooks.presetPhase2')}</option>
                    {Object.keys(presets).map((n) => <option key={n} value={n}>{n}</option>)}
                  </select>
                  {presetSel && !presetSel.startsWith('__') && (
                    <Button variant="subtle" onClick={deletePreset}>{t('common.delete')}</Button>
                  )}
                </div>
                <div className="mt-1.5 flex items-center gap-2">
                  <Input className="flex-1" placeholder={t('playbooks.presetName')} value={presetName} onChange={(e) => setPresetName(e.target.value)} />
                  <Button variant="subtle" disabled={!presetName.trim()} onClick={savePreset}>{t('common.save')}</Button>
                </div>
              </Field>
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
              <label className={`flex items-center gap-2 text-sm ${off.backupToUser ? 'text-[var(--text-dim)]' : 'text-[var(--text-faint)] opacity-60'}`}>
                <input type="checkbox" disabled={!off.backupToUser} checked={off.backupChats}
                  onChange={(e) => setOff({ ...off, backupChats: e.target.checked })} />
                {t('playbooks.backupChats')}{!off.backupToUser ? ` — ${t('playbooks.backupChatsNeedsTarget')}` : ''}
              </label>
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
                  <div className="font-medium text-[var(--text)]">{stepName(s)}{stepDetail(s) ? <span className="ml-2 text-xs text-[var(--text-faint)]">{stepDetail(s)}</span> : null}</div>
                  {/* Live percentage for the step in flight (e.g. the OneDrive backup). */}
                  {s.running && job?.progress && <div className="text-xs text-[var(--accent2)]">{job.progress}</div>}
                  {s.error && <div className="text-xs text-[var(--danger)]">{s.errorCode ? `${s.errorCode}: ` : ''}{s.error}</div>}
                  {s.hint && <div className="text-xs text-[var(--warn)]">{t('playbooks.permissionHint', { p: s.hint })}</div>}
                  {/* Known Graph limitation: Exchange-managed groups are not editable via Graph. */}
                  {s.error && /mail-enabled security|distribution list/i.test(s.error) && (
                    <div className="text-xs text-[var(--warn)]">{t('playbooks.exchangeGroupHint')}</div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </Page>
  )
}
