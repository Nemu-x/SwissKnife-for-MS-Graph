import { useEffect, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { UserPlus, UserMinus, Play, CheckCircle2, XCircle, Loader2, X, Trash2, Save } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { Button, Field, Input, Spinner, ErrorNote, Badge, Select } from '../components/ui'
import { MultiSelect, type Option } from '../components/MultiSelect'
import { EntityPicker } from '../components/EntityPicker'
import { UpnInput } from '../components/UpnInput'
import { useConfirm } from '../lib/useConfirm'
import { useStore, type PlaybookLiveStep } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import { skuFriendly } from '../lib/skuNames'
import { humanBytes } from '../lib/format'
import type { services } from '../../wailsjs/go/models'

// Onboarding options worth reusing between hires — the "role profile" a ticket
// really means when it says "same setup as the rest of the support team".
type OnboardProfile = {
  usageLocation: string
  skuIds: string[]
  groupIds: string[]
  teamIds: string[]
  channels: Record<string, string[]> // teamId → channelIds
}

const loadStore = <T,>(key: string): Record<string, T> => {
  // Persisted data is untrusted: accept only an object of objects.
  try {
    const raw = JSON.parse(localStorage.getItem(key) || '{}')
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {}
    const out: Record<string, T> = {}
    for (const [k, v] of Object.entries(raw)) {
      if (v && typeof v === 'object' && !Array.isArray(v)) out[k] = v as T
    }
    return out
  } catch { return {} }
}

export function PlaybooksPage() {
  const { t } = useTranslation()
  const { readOnly, connected, toast, jobs, startPlaybook, cancelPlaybook, clearJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()

  // The run lives in the store, so steps and the final report survive both
  // navigation and switching between the onboarding and offboarding tiles.
  const job = jobs.playbook
  const busy = !!job?.running
  const result = (job?.result as services.PlaybookResult | null) ?? null
  const steps: PlaybookLiveStep[] = result?.steps ?? job?.steps ?? []
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})

  // Backend emits stable keys + params; the English text is the fallback.
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
  // Channels accumulate across teams: picking a second team no longer discards
  // what was chosen in the first.
  const [channels, setChannels] = useState<Record<string, string[]>>({})
  const [chTeam, setChTeam] = useState('')
  const [chOpts, setChOpts] = useState<Option[]>([])
  const [chLoading, setChLoading] = useState(false)

  const remembered = (k: string) => localStorage.getItem('defaults.offboard.' + k) || ''
  const [off, setOff] = useState({
    upn: '', block: true, revokeSessions: true,
    oof: false, oofMessage: '', forwardTo: remembered('forwardTo'), hideFromGal: false,
    calendarTo: remembered('calendarTo'), removeFromGroups: false,
    removeAllLicenses: true, backupToUser: remembered('backupToUser'), backupFolder: remembered('backupFolder'),
    backupChats: false, intuneAction: '', removeMfaMethods: false, deleteRegisteredDevices: false,
    transferOwnershipTo: remembered('transferOwnershipTo'), cancelFutureEvents: false, delete: false,
  })

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
  const [presets, setPresets] = useState(() => loadStore<Partial<OffOptions>>('playbook.presets'))
  const [presetSel, setPresetSel] = useState('')
  const [presetName, setPresetName] = useState('')
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

  // Role profiles for onboarding.
  const [profiles, setProfiles] = useState(() => loadStore<OnboardProfile>('playbook.profiles'))
  const [profileSel, setProfileSel] = useState('')
  const [profileName, setProfileName] = useState('')
  const applyProfile = (name: string) => {
    setProfileSel(name)
    const p = profiles[name]
    if (!p) return
    setOn((o) => ({ ...o, usageLocation: p.usageLocation || o.usageLocation }))
    setSkuIds(p.skuIds || [])
    setGroupIds(p.groupIds || [])
    setTeamIds(p.teamIds || [])
    setChannels(p.channels || {})
  }
  const saveProfile = () => {
    const name = profileName.trim()
    if (!name) return
    const next = {
      ...profiles,
      [name]: { usageLocation: on.usageLocation, skuIds, groupIds, teamIds, channels } as OnboardProfile,
    }
    setProfiles(next)
    localStorage.setItem('playbook.profiles', JSON.stringify(next))
    setProfileSel(name); setProfileName('')
    toast('ok', t('playbooks.profileSaved', { name }))
  }
  const deleteProfile = () => {
    if (!profileSel) return
    const next = { ...profiles }
    delete next[profileSel]
    setProfiles(next)
    localStorage.setItem('playbook.profiles', JSON.stringify(next))
    setProfileSel('')
  }

  const [skuOpts, setSkuOpts] = useState<Option[]>([])
  const [groupOpts, setGroupOpts] = useState<Option[]>([])
  const [teamOpts, setTeamOpts] = useState<Option[]>([])
  const [loadingOpts, setLoadingOpts] = useState(false)

  useEffect(() => {
    if (!connected) return
    setLoadingOpts(true)
    Promise.allSettled([api.licensing.skus(), api.groups.list('', 0), api.teams.all()]).then(([s, g, tm]) => {
      const rows = (r: PromiseSettledResult<unknown>) => (r.status === 'fulfilled' ? (r.value as GraphObject[] | null) ?? [] : [])
      setSkuOpts(rows(s).map((x) => ({ value: x.skuId, label: skuFriendly(x.skuPartNumber), sub: `${x.consumedUnits ?? ''}/${x.prepaidUnits?.enabled ?? ''}` })))
      setGroupOpts(rows(g).map((x) => ({ value: x.id, label: x.displayName || x.id, sub: x.mail })))
      setTeamOpts(rows(tm).map((x) => ({ value: x.id, label: x.displayName || x.id })))
      setLoadingOpts(false)
    })
  }, [connected])

  // Channels of the team currently open in the picker.
  useEffect(() => {
    setChOpts([])
    if (!chTeam) return
    setChLoading(true)
    api.teams.channels(chTeam)
      .then((chs) => setChOpts((chs || []).map((c) => ({ value: c.id, label: c.displayName || c.id, sub: c.membershipType }))))
      .catch((e) => toast('err', errMessage(e)))
      .finally(() => setChLoading(false))
  }, [chTeam])

  const teamName = (id: string) => teamOpts.find((o) => o.value === id)?.label || id
  const channelRefs = Object.entries(channels).flatMap(([teamId, ids]) => ids.map((channelId) => ({ teamId, channelId })))
  const channelTeams = Object.entries(channels).filter(([, ids]) => ids.length > 0)

  const runOnboard = () => {
    startPlaybook('onboard', on.upn, () => api.playbooks.onboard({
      displayName: on.displayName, upn: on.upn, mailNickname: on.mailNickname, password: on.password,
      usageLocation: on.usageLocation, skuIds, groupIds, teamIds, channelRefs,
    })).then((r) => {
      if (r) setStatus((s) => ({ ...s, onboard: { ok: !!r.ok, text: on.upn, at: Date.now() } }))
    })
  }

  const runOffboard = (confirm: string) => () => {
    startPlaybook('offboard', off.upn, () => api.playbooks.offboard({ ...off, confirm })).then((r) => {
      if (r) setStatus((s) => ({ ...s, offboard: { ok: !!r.ok, text: off.upn, at: Date.now() } }))
      if (!r || r.canceled) return // failed or canceled run: keep the previous defaults
      for (const k of ['forwardTo', 'calendarTo', 'backupToUser', 'backupFolder', 'transferOwnershipTo'] as const) {
        if (off[k]) localStorage.setItem('defaults.offboard.' + k, off[k])
      }
    })
  }

  const cancelBar = busy && (
    <div className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs text-[var(--accent2)]">
      <Loader2 size={13} className="shrink-0 animate-spin" />
      <span className="min-w-0 flex-1 truncate">{job?.progress || t('playbooks.runningNote')}</span>
      <button onClick={cancelPlaybook} disabled={job?.canceled} className="shrink-0 text-[var(--danger)] hover:underline disabled:opacity-50">
        {job?.canceled ? t('common.canceling') : t('common.cancel')}
      </button>
    </div>
  )

  const actions: TaskAction[] = [
    {
      id: 'onboard', label: t('playbooks.tileOnboard'), hint: t('playbooks.hintOnboard'),
      icon: <UserPlus size={16} />, variant: 'primary', write: true,
      note: (
        <>
          <p>{t('playbooks.noteOnboard')}</p>
          {channelTeams.length > 0 && (
            <p className="text-[var(--text-dim)]">
              {t('playbooks.channelsPicked', { n: channelRefs.length, teams: channelTeams.length })}
            </p>
          )}
        </>
      ),
      panel: (
        <TaskForm>
          <Field label={t('playbooks.profile')} hint={t('playbooks.profileHint')}>
            <div className="flex items-center gap-2">
              <Select value={profileSel} onChange={(e) => applyProfile(e.target.value)} className="min-w-0 flex-1">
                <option value="">{t('playbooks.profilePick')}</option>
                {Object.keys(profiles).map((n) => <option key={n} value={n}>{n}</option>)}
              </Select>
              {profileSel && <Button variant="subtle" onClick={deleteProfile}><Trash2 size={14} /></Button>}
            </div>
            <div className="mt-1.5 flex items-center gap-2">
              <Input className="flex-1" placeholder={t('playbooks.profileName')} value={profileName} onChange={(e) => setProfileName(e.target.value)} />
              <Button variant="subtle" disabled={!profileName.trim()} onClick={saveProfile}><Save size={14} /> {t('common.save')}</Button>
            </div>
          </Field>

          <div className="my-1 border-t border-[var(--border)]" />
          <Field label={t('playbooks.displayName')}><Input value={on.displayName} onChange={(e) => setOn({ ...on, displayName: e.target.value })} /></Field>
          <Field label={t('playbooks.upn')}><UpnInput value={on.upn} onChange={(v) => setOn({ ...on, upn: v })} /></Field>
          <Field label={t('playbooks.mailNickname')}><Input value={on.mailNickname} onChange={(e) => setOn({ ...on, mailNickname: e.target.value })} /></Field>
          <Field label={t('playbooks.password')}><Input type="password" value={on.password} onChange={(e) => setOn({ ...on, password: e.target.value })} /></Field>
          <Field label={t('playbooks.usageLocation')}><Input value={on.usageLocation} onChange={(e) => setOn({ ...on, usageLocation: e.target.value })} placeholder="US" /></Field>

          <Field label={t('playbooks.licenses')}>
            <MultiSelect options={skuOpts} selected={skuIds} onChange={setSkuIds} loading={loadingOpts} placeholder={t('playbooks.pickLicenses')} />
          </Field>
          <Field label={t('playbooks.groups')}>
            <MultiSelect options={groupOpts} selected={groupIds} onChange={setGroupIds} loading={loadingOpts} placeholder={t('playbooks.pickGroups')} />
          </Field>
          <Field label={t('playbooks.teams')}>
            <MultiSelect options={teamOpts} selected={teamIds} onChange={setTeamIds} loading={loadingOpts} placeholder={t('playbooks.pickTeams')} />
          </Field>
          <Field label={t('playbooks.channels')} hint={t('playbooks.channelsHint')}>
            <div className="flex flex-col gap-2">
              <EntityPicker value={chTeam} onChange={setChTeam} load={async () => teamOpts} placeholder={t('playbooks.pickTeamForChannels')} />
              {chTeam && (
                <MultiSelect
                  options={chOpts}
                  selected={channels[chTeam] || []}
                  onChange={(ids) => setChannels({ ...channels, [chTeam]: ids })}
                  loading={chLoading}
                  placeholder={t('playbooks.pickChannels')}
                />
              )}
              {channelTeams.length > 0 && (
                <div className="flex flex-wrap gap-1.5">
                  {channelTeams.map(([teamId, ids]) => (
                    <span key={teamId} className="flex items-center gap-1 rounded-md bg-[var(--bg-elev-2)] px-2 py-0.5 text-xs text-[var(--text-dim)]">
                      {teamName(teamId)} · {ids.length}
                      <button onClick={() => setChannels({ ...channels, [teamId]: [] })} className="text-[var(--text-faint)] hover:text-[var(--danger)]">
                        <X size={11} />
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          </Field>

          <Button variant="primary" disabled={readOnly || busy || !on.upn || !on.displayName || !on.mailNickname || !on.password} onClick={runOnboard}>
            {busy ? <Spinner /> : <Play size={15} />} {t('playbooks.run')}
          </Button>
          {cancelBar}
        </TaskForm>
      ),
    },
    {
      id: 'offboard', label: t('playbooks.tileOffboard'), hint: t('playbooks.hintOffboard'),
      icon: <UserMinus size={16} />, variant: 'danger', write: true,
      note: (
        <>
          <p>{t('playbooks.noteOffboard')}</p>
          {off.removeAllLicenses && <p className="text-[var(--warn)]">{t('playbooks.removeLicensesWarn')}</p>}
          {off.intuneAction === 'wipe' && <p className="text-[var(--danger)]">{t('playbooks.intuneWipeWarn')}</p>}
          {off.delete && <p className="text-[var(--danger)]">{t('playbooks.deleteUserWarn')}</p>}
        </>
      ),
      panel: (
        <TaskForm>
          <Field label={t('playbooks.presets')}>
            <div className="flex items-center gap-2">
              <Select value={presetSel} onChange={(e) => applyPreset(e.target.value)} className="min-w-0 flex-1">
                <option value="">{t('playbooks.presetPick')}</option>
                <option value="__phase1">{t('playbooks.presetPhase1')}</option>
                <option value="__phase2">{t('playbooks.presetPhase2')}</option>
                {Object.keys(presets).map((n) => <option key={n} value={n}>{n}</option>)}
              </Select>
              {presetSel && !presetSel.startsWith('__') && (
                <Button variant="subtle" onClick={deletePreset}><Trash2 size={14} /></Button>
              )}
            </div>
            <div className="mt-1.5 flex items-center gap-2">
              <Input className="flex-1" placeholder={t('playbooks.presetName')} value={presetName} onChange={(e) => setPresetName(e.target.value)} />
              <Button variant="subtle" disabled={!presetName.trim()} onClick={savePreset}><Save size={14} /> {t('common.save')}</Button>
            </div>
          </Field>

          <div className="my-1 border-t border-[var(--border)]" />
          <Field label={t('playbooks.upn')}><UpnInput value={off.upn} onChange={(v) => setOff({ ...off, upn: v })} /></Field>

          {([['block', 'block'], ['revokeSessions', 'revoke'], ['removeMfaMethods', 'removeMfa'], ['oof', 'oof'], ['hideFromGal', 'hideFromGal'], ['cancelFutureEvents', 'cancelEvents'], ['removeFromGroups', 'removeFromGroups'], ['deleteRegisteredDevices', 'deleteRegisteredDevices'], ['removeAllLicenses', 'removeLicenses'], ['delete', 'deleteUser']] as const).map(([k, label]) => (
            <label key={k} className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
              <input type="checkbox" checked={(off as any)[k]} onChange={(e) => setOff({ ...off, [k]: e.target.checked })} />
              <span className={k === 'delete' ? 'text-[var(--danger)]' : ''}>{t(`playbooks.${label}`)}</span>
            </label>
          ))}
          {off.oof && (
            <Field label={t('playbooks.oofMessage')}>
              <Input value={off.oofMessage} onChange={(e) => setOff({ ...off, oofMessage: e.target.value })} />
            </Field>
          )}

          <Field label={t('playbooks.intuneAction')}>
            <Select value={off.intuneAction} onChange={(e) => setOff({ ...off, intuneAction: e.target.value })} className="w-full">
              <option value="">{t('playbooks.intuneNone')}</option>
              <option value="retire">{t('playbooks.intuneRetire')}</option>
              <option value="wipe">{t('playbooks.intuneWipe')}</option>
            </Select>
          </Field>

          <Field label={t('playbooks.transferOwnershipTo')} hint={t('playbooks.transferOwnershipHint')}>
            <UpnInput value={off.transferOwnershipTo} onChange={(v) => setOff({ ...off, transferOwnershipTo: v })} placeholder="lead@contoso.com" />
          </Field>
          <Field label={t('playbooks.forwardTo')}><UpnInput value={off.forwardTo} onChange={(v) => setOff({ ...off, forwardTo: v })} placeholder="manager@contoso.com" /></Field>
          <Field label={t('playbooks.calendarTo')}><UpnInput value={off.calendarTo} onChange={(v) => setOff({ ...off, calendarTo: v })} placeholder="manager@contoso.com" /></Field>
          <Field label={t('playbooks.backupTo')}><UpnInput value={off.backupToUser} onChange={(v) => setOff({ ...off, backupToUser: v })} placeholder="backup@contoso.com" /></Field>
          <Field label={t('playbooks.backupFolder')}><Input value={off.backupFolder} onChange={(e) => setOff({ ...off, backupFolder: e.target.value })} /></Field>
          <label className={`flex items-center gap-2 text-sm ${off.backupToUser ? 'text-[var(--text-dim)]' : 'text-[var(--text-faint)] opacity-60'}`}>
            {/* Display mirrors the backend guard: no target — no chat backup. */}
            <input type="checkbox" disabled={!off.backupToUser} checked={off.backupChats && !!off.backupToUser}
              onChange={(e) => setOff({ ...off, backupChats: e.target.checked })} />
            {t('playbooks.backupChats')}{!off.backupToUser ? ` — ${t('playbooks.backupChatsNeedsTarget')}` : ''}
          </label>

          <Button variant="danger" disabled={readOnly || busy || !off.upn} onClick={() => askConfirm(off.upn, (c) => runOffboard(c)())}>
            {busy ? <Spinner /> : <Play size={15} />} {t('playbooks.run')}
          </Button>
          {cancelBar}
        </TaskForm>
      ),
    },
  ]

  const report = (
    <div className="flex h-full flex-col">
      <div className="flex items-center justify-between gap-2 border-b border-[var(--border)] px-3 py-2">
        <span className="text-sm font-medium">{t('playbooks.steps')}</span>
        <div className="flex items-center gap-2">
          {result?.canceled && <Badge kind="warn">{t('common.canceled')}</Badge>}
          {!busy && steps.length > 0 && (
            <Button variant="ghost" className="!px-2 !py-1 text-xs" onClick={() => clearJob('playbook')}>{t('common.clear')}</Button>
          )}
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-auto p-3">
        {job?.error && <ErrorNote>{job.error}</ErrorNote>}
        {steps.length === 0 && !busy && !job?.error && (
          <p className="text-sm text-[var(--text-faint)]">{t('playbooks.reportEmpty')}</p>
        )}
        <div className="flex flex-col gap-2">
          {steps.map((s, i) => (
            <div key={i} className="flex items-start gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-sm">
              {s.running
                ? <Loader2 size={16} className="mt-0.5 shrink-0 animate-spin text-[var(--accent2)]" />
                : s.ok
                  ? <CheckCircle2 size={16} className="mt-0.5 shrink-0 text-[var(--ok)]" />
                  : <XCircle size={16} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
              <div className="min-w-0">
                <div className="font-medium text-[var(--text)]">
                  {stepName(s)}
                  {stepDetail(s) ? <span className="ml-2 text-xs text-[var(--text-faint)]">{stepDetail(s)}</span> : null}
                </div>
                {s.running && job?.progress && <div className="text-xs text-[var(--accent2)]">{job.progress}</div>}
                {s.error && <div className="text-xs text-[var(--danger)]">{s.errorCode ? `${s.errorCode}: ` : ''}{s.error}</div>}
                {s.hint && <div className="text-xs text-[var(--warn)]">{t('playbooks.permissionHint', { p: s.hint })}</div>}
                {s.error && /mail-enabled security|distribution list/i.test(s.error) && (
                  <div className="text-xs text-[var(--warn)]">{t('playbooks.exchangeGroupHint')}</div>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="playbooks"
        title={t('playbooks.title')}
        subtitle={t('playbooks.subtitle')}
        actions={actions}
        status={status}
        busy={busy}
        busyLabel={job?.progress || t('playbooks.running')}
        hasResult={steps.length > 0 || busy || !!job?.error}
        onClearResult={() => clearJob('playbook')}
        result={report}
      />
    </>
  )
}
