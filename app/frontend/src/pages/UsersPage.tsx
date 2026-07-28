import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import {
  UserCog, Ban, KeyRound, LogOut, UserPlus, Trash2, UserSquare, ShieldAlert,
  ListRestart, RotateCcw, Ticket, MailPlus, Copy, FileJson, Search, ArrowLeftRight, CopyPlus,
} from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction, type ActionStatus } from '../components/TaskPage'
import { ResultView } from '../components/ResultView'
import { Button, Field, Input, Textarea, Spinner } from '../components/ui'
import { UpnInput } from '../components/UpnInput'
import { EntityPicker } from '../components/EntityPicker'
import { loadUsers } from '../lib/pickers'
import { useAsync } from '../lib/useAsync'
import { useConfirm } from '../lib/useConfirm'
import { useStore } from '../lib/store'
import { skuFriendly } from '../lib/skuNames'
import { api, errMessage, type GraphObject } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

// Access kinds a mirror run can copy, in the order the tiles offer them.
const MIRROR_KINDS = ['group', 'team', 'channel', 'role', 'license'] as const
const STATUS_RANK: Record<string, number> = { missing: 0, both: 1, targetOnly: 2 }

export function UsersPage() {
  const { t } = useTranslation()
  const { readOnly, toast, cache, setCache, jobs, patchJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const res = useAsync<GraphObject[] | GraphObject>()

  const [search, setSearch] = useState('')
  const [target, setTarget] = useState('')
  const [pw, setPw] = useState('')
  const [force, setForce] = useState(true)
  const [create, setCreate] = useState({ name: '', upn: '', nick: '', pw: '', loc: '' })
  const [mgr, setMgr] = useState('')
  const [loc, setLoc] = useState('')
  const [restoreId, setRestoreId] = useState('')
  const [patch, setPatch] = useState('{\n  "jobTitle": "",\n  "department": ""\n}')
  // TAP is a one-time secret shown once — back it with the store cache so
  // navigating away does not destroy it before the operator copies it.
  const [tap, setTapLocal] = useState<GraphObject | null>(() => cache['users.tap'] ?? null)
  const setTap = (v: GraphObject | null) => { setTapLocal(v); setCache('users.tap', v) }
  const [tapQr, setTapQrLocal] = useState<string>(() => cache['users.tapQr'] ?? '')
  const setTapQr = (v: string) => { setTapQrLocal(v); setCache('users.tapQr', v) }
  const closeTap = () => { setTap(null); setTapQr('') }
  const [tapLifetime, setTapLifetime] = useState(60)
  const [tapOnce, setTapOnce] = useState(true)
  const [invite, setInvite] = useState({ email: '', name: '', sendMail: true })
  const [status, setStatus] = useState<Record<string, ActionStatus>>({})
  // Mirror: "give the target the same access this source user has".
  const [mirrorSource, setMirrorSource] = useState('')
  const [mirrorKinds, setMirrorKinds] = useState<Record<string, boolean>>({
    group: true, team: true, channel: true, role: false, license: false,
  })
  // The comparison is kept raw so the pane can be re-rendered when the operator
  // switches between "only what will be copied" and the full picture.
  const [diff, setDiff] = useState<services.AccessRow[] | null>(null)
  const [showAll, setShowAll] = useState(false)

  const mark = (id: string, ok: boolean, text: string) =>
    setStatus((s) => ({ ...s, [id]: { ok, text, at: Date.now() } }))

  const listUsers = () => res.run(() => api.users.list(search, 0))
  const doWrite = async (id: string, fn: () => Promise<any>, ok: string) => {
    try { await fn(); mark(id, true, ok); toast('ok', ok) }
    catch (e) { const m = errMessage(e); mark(id, false, m); toast('err', m) }
  }
  const doShow = async (id: string, fn: () => Promise<any>, ok: string) => {
    try { res.setData(await fn()); mark(id, true, ok); toast('ok', ok) }
    catch (e) { const m = errMessage(e); mark(id, false, m); toast('err', m) }
  }

  // Issue a Temporary Access Pass and show it once, with a QR for phone entry.
  const makeTap = async () => {
    try {
      const r = await api.authMethods.createTAP(target, tapLifetime, tapOnce)
      setTap(r)
      setTapQr('')
      mark('tap', true, t('users.createTap'))
      const QRCode = await import('qrcode')
      setTapQr(await QRCode.toDataURL(String(r.temporaryAccessPass || ''), { margin: 1, width: 220 }))
    } catch (e) { const m = errMessage(e); mark('tap', false, m); toast('err', m) }
  }

  // The diff and the copy report are Graph shapes: translate them here so the
  // results view shows sentences instead of keys.
  const counts = (rows: services.AccessRow[]) => ({
    copy: rows.filter((r) => r.status === 'missing' && r.copyable).length,
    blocked: rows.filter((r) => r.status === 'missing' && !r.copyable).length,
    both: rows.filter((r) => r.status === 'both').length,
    targetOnly: rows.filter((r) => r.status === 'targetOnly').length,
  })

  // A 60-row dump answers nothing. By default the pane shows only what the copy
  // would actually touch; the rest is one checkbox away.
  const diffRows = (rows: services.AccessRow[], all: boolean) =>
    [...rows]
      .filter((r) => all || r.status === 'missing')
      .sort((a, b) => (STATUS_RANK[a.status] ?? 9) - (STATUS_RANK[b.status] ?? 9) || a.kind.localeCompare(b.kind))
      // Name first: the results list labels each row by its first scalar field.
      .map((r) => ({
        [t('mirror.colName')]: `${r.status === 'missing' ? (r.copyable ? '+ ' : '! ') : r.status === 'both' ? '= ' : '· '}${r.kind === 'license' ? skuFriendly(r.name) : r.name}`,
        [t('mirror.colWhat')]: t(`mirror.kind.${r.kind}`),
        [t('mirror.colTeam')]: r.teamName || '—',
        [t('mirror.colStatus')]: t(`mirror.status.${r.status}`),
        [t('mirror.colCopyable')]: r.status !== 'missing' ? '—' : r.copyable ? t('mirror.yes') : t('mirror.no'),
        [t('mirror.colNote')]: r.reasonKey ? t(r.reasonKey) : '',
      }))

  const stepRows = (result: services.PlaybookResult) =>
    (result.steps || []).map((s) => ({
      [t('mirror.colName')]: s.detail || '',
      [t('mirror.colStep')]: s.nameKey ? t(s.nameKey, { defaultValue: s.name }) : s.name,
      [t('mirror.colResult')]: s.ok ? '✓' : '✗',
      [t('mirror.colNote')]: s.ok ? '' : s.detailKey ? t(s.detailKey) : s.error || '',
    }))

  // The scan runs in the store's job slot, so its live line survives navigation
  // and a second click cannot start a competing run.
  const mirrorJob = jobs.mirror
  const mirrorBusy = !!mirrorJob?.running

  const compareAccess = async () => {
    if (mirrorBusy) return
    patchJob('mirror', { running: true, error: null, progress: t('mirror.starting'), startedAt: Date.now() })
    try {
      const rows = await api.mirror.compare(mirrorSource, target)
      setDiff(rows)
      res.setData(diffRows(rows, showAll) as any)
      const c = counts(rows)
      mark('compareAccess', true, t('mirror.summary', c))
    } catch (e) { const m = errMessage(e); mark('compareAccess', false, m); toast('err', m) }
    finally { patchJob('mirror', { running: false, progress: '' }) }
  }

  const copyAccess = (confirm: string) => async () => {
    if (mirrorBusy) return
    const kinds = MIRROR_KINDS.filter((k) => mirrorKinds[k])
    patchJob('mirror', { running: true, error: null, progress: t('mirror.starting'), startedAt: Date.now() })
    try {
      const r = await api.mirror.copy(mirrorSource, target, kinds, confirm)
      setDiff(null)
      res.setData(stepRows(r) as any)
      const failed = (r.steps || []).filter((s) => !s.ok).length
      mark('copyAccess', failed === 0, failed === 0 ? t('mirror.copied') : t('mirror.copiedWithErrors', { n: failed }))
      toast(failed === 0 ? 'ok' : 'err', failed === 0 ? t('mirror.copied') : t('mirror.copiedWithErrors', { n: failed }))
    } catch (e) { const m = errMessage(e); mark('copyAccess', false, m); toast('err', m) }
    finally { patchJob('mirror', { running: false, progress: '' }) }
  }

  // Live line + cancel, shown inside both mirror tiles while a scan runs.
  const mirrorProgress = mirrorBusy && (
    <div className="flex items-center gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2 text-xs text-[var(--accent2)]">
      <Spinner />
      <span className="min-w-0 flex-1 truncate">{mirrorJob?.progress || t('mirror.starting')}</span>
      <button onClick={() => api.mirror.cancel()} className="shrink-0 text-[var(--danger)] hover:underline">
        {t('common.cancel')}
      </button>
    </div>
  )

  const diffToggle = diff && (
    <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
      <input type="checkbox" checked={showAll}
        onChange={(e) => { setShowAll(e.target.checked); res.setData(diffRows(diff, e.target.checked) as any) }} />
      {t('mirror.showAll', { n: diff.length })}
    </label>
  )

  const kindChecks = (
    <div className="flex flex-col gap-1.5">
      {MIRROR_KINDS.map((k) => (
        <label key={k} className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
          <input type="checkbox" checked={!!mirrorKinds[k]}
            onChange={(e) => setMirrorKinds({ ...mirrorKinds, [k]: e.target.checked })} />
          {t(`mirror.kind.${k}`)}
        </label>
      ))}
    </div>
  )

  const targetField = (
    <Field label={t('common.user')}>
      <EntityPicker value={target} onChange={setTarget} load={loadUsers} placeholder={t('users.pickUser')} />
    </Field>
  )

  const actions: TaskAction[] = [
    {
      id: 'list', label: t('users.tileList'), hint: t('users.hintList'), icon: <Search size={16} />,
      onClick: listUsers,
    },
    {
      id: 'snapshot', label: t('users.tileSnapshot'), hint: t('users.hintSnapshot'), icon: <UserCog size={16} />, variant: 'primary',
      note: <p>{t('users.noteSnapshot')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <Button variant="primary" disabled={!target} onClick={() => res.run(() => api.users.snapshot(target) as any)}>
            <UserCog size={15} /> {t('users.snapshot')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'compareAccess', label: t('mirror.tileCompare'), hint: t('mirror.hintCompare'), icon: <ArrowLeftRight size={16} />, variant: 'primary',
      note: <p>{t('mirror.noteCompare')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('mirror.source')}>
            <EntityPicker value={mirrorSource} onChange={setMirrorSource} load={loadUsers} placeholder={t('users.pickUser')} />
          </Field>
          <Field label={t('mirror.target')}>
            <EntityPicker value={target} onChange={setTarget} load={loadUsers} placeholder={t('users.pickUser')} />
          </Field>
          {diffToggle}
          <Button variant="primary" disabled={mirrorBusy || !mirrorSource || !target || mirrorSource === target} onClick={compareAccess}>
            {mirrorBusy ? <Spinner /> : <ArrowLeftRight size={15} />} {t('mirror.compare')}
          </Button>
          {mirrorProgress}
        </TaskForm>
      ),
    },
    {
      id: 'copyAccess', label: t('mirror.tileCopy'), hint: t('mirror.hintCopy'), icon: <CopyPlus size={16} />, variant: 'primary', write: true,
      note: (
        <>
          <p>{t('mirror.noteCopy')}</p>
          <p className="text-[var(--warn)]">{t('mirror.noteCopyWarn')}</p>
        </>
      ),
      panel: (
        <TaskForm>
          <Field label={t('mirror.source')}>
            <EntityPicker value={mirrorSource} onChange={setMirrorSource} load={loadUsers} placeholder={t('users.pickUser')} />
          </Field>
          <Field label={t('mirror.target')}>
            <EntityPicker value={target} onChange={setTarget} load={loadUsers} placeholder={t('users.pickUser')} />
          </Field>
          <Field label={t('mirror.whatToCopy')}>{kindChecks}</Field>
          <Button variant="subtle" disabled={mirrorBusy || !mirrorSource || !target || mirrorSource === target} onClick={compareAccess}>
            {mirrorBusy ? <Spinner /> : <ArrowLeftRight size={15} />} {t('mirror.previewFirst')}
          </Button>
          {diffToggle}
          <Button variant="primary"
            disabled={readOnly || mirrorBusy || !mirrorSource || !target || mirrorSource === target || !MIRROR_KINDS.some((k) => mirrorKinds[k])}
            onClick={() => askConfirm(target, (c) => copyAccess(c)())}>
            <CopyPlus size={15} /> {t('mirror.copy')}
          </Button>
          {mirrorProgress}
        </TaskForm>
      ),
    },
    {
      id: 'block', label: t('users.tileBlock'), hint: t('users.hintBlock'), icon: <Ban size={16} />, write: true,
      panel: (
        <TaskForm>
          {targetField}
          <div className="grid grid-cols-2 gap-2">
            <Button variant="primary" disabled={readOnly || !target} onClick={() => doWrite('block', () => api.users.block(target), t('users.block'))}>
              <Ban size={15} /> {t('users.block')}
            </Button>
            <Button variant="subtle" disabled={readOnly || !target} onClick={() => doWrite('block', () => api.users.unblock(target), t('users.unblock'))}>
              {t('users.unblock')}
            </Button>
          </div>
        </TaskForm>
      ),
    },
    {
      id: 'sessions', label: t('users.tileSessions'), hint: t('users.hintSessions'), icon: <LogOut size={16} />, write: true,
      note: <p>{t('users.noteSessions')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <Button variant="danger" disabled={readOnly || !target}
            onClick={() => askConfirm(target, (c) => doWrite('sessions', () => api.users.revokeSessions(target, c), t('users.revokeSessions')))}>
            <LogOut size={15} /> {t('users.revokeSessions')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'password', label: t('users.tilePassword'), hint: t('users.hintPassword'), icon: <KeyRound size={16} />, write: true,
      panel: (
        <TaskForm>
          {targetField}
          <Field label={t('users.newPassword')}><Input type="password" value={pw} onChange={(e) => setPw(e.target.value)} /></Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={force} onChange={(e) => setForce(e.target.checked)} /> {t('users.forceChange')}
          </label>
          <Button variant="danger" disabled={readOnly || !pw || !target}
            onClick={() => askConfirm(target, (c) => doWrite('password', () => api.users.resetPassword(target, pw, force, c), t('users.resetPassword')))}>
            <KeyRound size={15} /> {t('users.resetPassword')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'mfa', label: t('users.tileMfa'), hint: t('users.hintMfa'), icon: <ShieldAlert size={16} />, write: true,
      note: <p>{t('users.noteMfa')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <Button variant="subtle" disabled={!target} onClick={() => res.run(() => api.authMethods.list(target))}>
            {t('users.authMethods')}
          </Button>
          <Button variant="danger" disabled={readOnly || !target}
            onClick={() => askConfirm(target, (c) => doShow('mfa', () => api.authMethods.resetMFA(target, c), t('users.resetMfa')))}>
            <RotateCcw size={15} /> {t('users.resetMfa')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'tap', label: t('users.tileTap'), hint: t('users.hintTap'), icon: <Ticket size={16} />, write: true,
      note: <p>{t('users.tapHint')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <div className="flex items-end gap-3">
            <Field label={t('users.tapLifetime')}>
              <Input type="number" value={tapLifetime} onChange={(e) => setTapLifetime(Math.max(10, Number(e.target.value) || 60))} className="w-24" />
            </Field>
            <label className="flex items-center gap-2 pb-1.5 text-sm text-[var(--text-dim)]">
              <input type="checkbox" checked={tapOnce} onChange={(e) => setTapOnce(e.target.checked)} /> {t('users.tapOnce')}
            </label>
          </div>
          <Button variant="primary" disabled={readOnly || !target} onClick={makeTap}>
            <Ticket size={15} /> {t('users.createTap')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'manager', label: t('users.tileManager'), hint: t('users.hintManager'), icon: <UserSquare size={16} />, write: true,
      panel: (
        <TaskForm>
          {targetField}
          <Button variant="subtle" disabled={!target} onClick={() => res.run(() => api.users.getManager(target))}>
            <UserSquare size={15} /> {t('users.getManager')}
          </Button>
          <Field label={t('users.manager')}><UpnInput value={mgr} onChange={setMgr} placeholder="manager@contoso.com" /></Field>
          <Button variant="primary" disabled={readOnly || !target || !mgr} onClick={() => doWrite('manager', () => api.users.setManager(target, mgr), t('users.setManager'))}>
            {t('users.setManager')}
          </Button>
          <Field label={t('users.usageLocation')}><Input value={loc} onChange={(e) => setLoc(e.target.value)} placeholder="US" /></Field>
          <Button variant="subtle" disabled={readOnly || !target || !loc} onClick={() => doWrite('manager', () => api.users.setUsageLocation(target, loc), t('users.setUsageLocation'))}>
            {t('users.setUsageLocation')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'patch', label: t('users.tilePatch'), hint: t('users.hintPatch'), icon: <FileJson size={16} />, write: true,
      note: <p>{t('users.notePatch')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <Field label={t('users.updateFields')}><Textarea rows={5} value={patch} onChange={(e) => setPatch(e.target.value)} /></Field>
          <Button variant="primary" disabled={readOnly || !target} onClick={() => doWrite('patch', () => api.users.update(target, patch), t('users.update'))}>
            {t('users.update')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'create', label: t('users.tileCreate'), hint: t('users.hintCreate'), icon: <UserPlus size={16} />, variant: 'primary', write: true,
      note: <p>{t('users.noteCreate')}</p>,
      panel: (
        <TaskForm>
          <Field label={t('users.displayName')}><Input value={create.name} onChange={(e) => setCreate({ ...create, name: e.target.value })} /></Field>
          <Field label={t('users.upn')}><UpnInput value={create.upn} onChange={(v) => setCreate({ ...create, upn: v })} /></Field>
          <Field label={t('users.mailNickname')}><Input value={create.nick} onChange={(e) => setCreate({ ...create, nick: e.target.value })} /></Field>
          <Field label={t('users.newPassword')}><Input type="password" value={create.pw} onChange={(e) => setCreate({ ...create, pw: e.target.value })} /></Field>
          <Field label={t('users.usageLocation')}><Input value={create.loc} onChange={(e) => setCreate({ ...create, loc: e.target.value })} placeholder="US" /></Field>
          <Button variant="primary" disabled={readOnly || !create.name || !create.upn || !create.nick || !create.pw}
            onClick={() => doShow('create', () => api.users.create(create.name, create.upn, create.nick, create.pw, true, create.loc), t('users.createUser'))}>
            <UserPlus size={15} /> {t('users.createUser')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'invite', label: t('users.tileInvite'), hint: t('users.hintInvite'), icon: <MailPlus size={16} />, write: true,
      panel: (
        <TaskForm>
          <Field label={t('users.inviteEmail')}><Input placeholder="guest@example.com" value={invite.email} onChange={(e) => setInvite({ ...invite, email: e.target.value })} /></Field>
          <Field label={t('users.inviteName')}><Input value={invite.name} onChange={(e) => setInvite({ ...invite, name: e.target.value })} /></Field>
          <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
            <input type="checkbox" checked={invite.sendMail} onChange={(e) => setInvite({ ...invite, sendMail: e.target.checked })} /> {t('users.inviteSendMail')}
          </label>
          <Button variant="primary" disabled={readOnly || !invite.email}
            onClick={() => doShow('invite', () => api.users.inviteGuest(invite.email, invite.name, '', '', invite.sendMail), t('users.inviteGuest'))}>
            <MailPlus size={15} /> {t('users.inviteGuest')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'delete', label: t('users.tileDelete'), hint: t('users.hintDelete'), icon: <Trash2 size={16} />, variant: 'danger', write: true,
      note: <p className="text-[var(--danger)]">{t('users.noteDelete')}</p>,
      panel: (
        <TaskForm>
          {targetField}
          <Button variant="danger" disabled={readOnly || !target}
            onClick={() => askConfirm(target, (c) => doWrite('delete', () => api.users.delete(target, c), t('users.deleteUser')))}>
            <Trash2 size={15} /> {t('users.deleteUser')}
          </Button>
        </TaskForm>
      ),
    },
    {
      id: 'restore', label: t('users.tileRestore'), hint: t('users.hintRestore'), icon: <ListRestart size={16} />, write: true,
      note: <p>{t('users.noteRestore')}</p>,
      panel: (
        <TaskForm>
          <Button variant="subtle" onClick={() => res.run(() => api.users.listDeleted(0))}>
            <ListRestart size={15} /> {t('users.listDeleted')}
          </Button>
          <Field label={t('users.objectId')}><Input value={restoreId} placeholder="object id" onChange={(e) => setRestoreId(e.target.value)} /></Field>
          <Button variant="primary" disabled={readOnly || !restoreId} onClick={() => doShow('restore', () => api.users.restoreDeleted(restoreId), t('users.restore'))}>
            <RotateCcw size={15} /> {t('users.restore')}
          </Button>
        </TaskForm>
      ),
    },
  ]

  return (
    <>
      {confirmElement}
      {tap && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60" onClick={closeTap}>
          <div className="w-[380px] rounded-2xl border border-[var(--border)] bg-[var(--bg-elev)] p-5" onClick={(e) => e.stopPropagation()}>
            <div className="mb-2 text-sm font-medium">{t('users.tapTitle')}</div>
            <div className="rounded-lg bg-[var(--bg)] p-3 text-center font-mono text-xl tracking-wider">{String(tap.temporaryAccessPass || '')}</div>
            {tapQr && <img src={tapQr} alt="TAP QR" width={220} height={220} className="mx-auto mt-3 rounded-lg bg-white p-2" />}
            <div className="mt-2 text-center text-xs text-[var(--text-faint)]">{t('users.tapHint')}</div>
            <div className="mt-3 flex gap-2">
              <Button variant="primary" className="flex-1"
                onClick={() => { navigator.clipboard.writeText(String(tap.temporaryAccessPass || '')); toast('ok', t('users.tapCopied')) }}>
                <Copy size={15} /> {t('users.tapCopy')}
              </Button>
              <Button className="flex-1" onClick={closeTap}>{t('users.tapClose')}</Button>
            </div>
          </div>
        </div>
      )}
      <TaskPage
        pageId="users"
        title={t('nav.users')}
        subtitle={t('users.subtitle')}
        search={{ value: search, onChange: setSearch, onSubmit: listUsers, placeholder: t('users.searchHint') }}
        actions={actions}
        status={status}
        busy={res.loading || mirrorBusy}
        busyLabel={mirrorBusy ? mirrorJob?.progress || t('mirror.starting') : undefined}
        onClearResult={() => { res.reset(); setDiff(null); setShowAll(false) }}
        hasResult={!!res.data || res.loading || !!res.error}
        result={<ResultView data={res.data} loading={res.loading} error={res.error} onUseId={setTarget} />}
      />
    </>
  )
}
