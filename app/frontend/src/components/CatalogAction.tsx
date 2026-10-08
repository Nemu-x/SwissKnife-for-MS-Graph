import { useEffect, useRef, useState, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { Ban, LogOut, UserPlus, Plus, Zap, UserSquare, Globe, Send, FolderLock, KeySquare, AtSign, Inbox, Forward, CalendarCheck, ListPlus, ShieldOff, ScrollText, UsersRound, ListChecks, BadgeCheck, Play, MailX, Siren, Flag, ShieldBan, Archive, ArchiveRestore, Eye, Check, ArrowRight, Minus } from 'lucide-react'
import { Button, Field, Input, Spinner } from './ui'
import { EntityPicker } from './EntityPicker'
import type { TaskAction } from './TaskPage'
import { useStore } from '../lib/store'
import { useConfirm } from '../lib/useConfirm'
import { useTaskStatus } from '../lib/useTaskStatus'
import { loadGroups, loadSkus, loadUsers } from '../lib/pickers'
import { skuFriendly } from '../lib/skuNames'
import { api, errMessage, errParsed } from '../lib/api'
import type { engine } from '../../wailsjs/go/models'

// Catalog actions (ADR-008) rendered as ordinary page tiles. A page places one
// with catalogTile(id) where the tile belongs; catalog actions for the page
// that no page placed (packs, new backends) are appended after the page's own
// tiles. The form is generated from the manifest, every write goes through a
// preview of what will change, and Apply runs exactly that preview.

export type CatalogEntry = engine.CatalogEntry

// Placeholder a page puts into its tile list; TaskPage swaps in the real tile.
export const catalogTile = (id: string): TaskAction => ({ id, label: '', catalog: true })

// Built-in actions reuse the wording of the tiles they replaced; anything else
// falls back to actions.<id>.label / .hint.
const UI: Record<string, { label: string; hint?: string; notes?: string[]; warn?: string[]; icon: ReactNode }> = {
  'user.signIn': { label: 'users.tileBlock', hint: 'users.hintBlock', icon: <Ban size={16} /> },
  'user.revokeSessions': { label: 'users.tileSessions', hint: 'users.hintSessions', notes: ['users.noteSessions'], icon: <LogOut size={16} /> },
  'user.manager': { label: 'actions.user.manager.label', hint: 'actions.user.manager.hint', icon: <UserSquare size={16} /> },
  'user.usageLocation': { label: 'actions.user.usageLocation.label', hint: 'actions.user.usageLocation.hint', notes: ['actions.user.usageLocation.note'], icon: <Globe size={16} /> },
  'mailbox.sendOnBehalf': { label: 'actions.mailbox.sendOnBehalf.label', hint: 'actions.mailbox.sendOnBehalf.hint', notes: ['actions.mailbox.sendOnBehalf.note'], icon: <Send size={16} /> },
  'mailbox.folderPermission': { label: 'actions.mailbox.folderPermission.label', hint: 'actions.mailbox.folderPermission.hint', notes: ['actions.mailbox.folderPermission.note'], icon: <FolderLock size={16} /> },
  'mailbox.fullAccess': { label: 'actions.mailbox.fullAccess.label', hint: 'actions.mailbox.fullAccess.hint', notes: ['actions.mailbox.fullAccess.note'], icon: <KeySquare size={16} /> },
  'mailbox.sendAs': { label: 'actions.mailbox.sendAs.label', hint: 'actions.mailbox.sendAs.hint', icon: <AtSign size={16} /> },
  'mailbox.type': { label: 'actions.mailbox.type.label', hint: 'actions.mailbox.type.hint', warn: ['actions.mailbox.type.note'], icon: <Inbox size={16} /> },
  'mailbox.forwarding': { label: 'actions.mailbox.forwarding.label', hint: 'actions.mailbox.forwarding.hint', notes: ['actions.mailbox.forwarding.note'], icon: <Forward size={16} /> },
  'mailbox.address': { label: 'actions.mailbox.address.label', hint: 'actions.mailbox.address.hint', icon: <AtSign size={16} /> },
  'mailbox.calendarProcessing': { label: 'actions.mailbox.calendarProcessing.label', hint: 'actions.mailbox.calendarProcessing.hint', notes: ['actions.mailbox.calendarProcessing.note'], icon: <CalendarCheck size={16} /> },
  'distributionList.membership': { label: 'actions.distributionList.membership.label', hint: 'actions.distributionList.membership.hint', notes: ['actions.distributionList.membership.note'], icon: <ListPlus size={16} /> },
  'transportRule.state': { label: 'actions.transportRule.state.label', hint: 'actions.transportRule.state.hint', warn: ['actions.transportRule.state.note'], icon: <ShieldOff size={16} /> },
  'teams.userPolicy': { label: 'actions.teams.userPolicy.label', hint: 'actions.teams.userPolicy.hint', notes: ['actions.teams.userPolicy.note'], icon: <ScrollText size={16} /> },
  'teams.groupPolicy': { label: 'actions.teams.groupPolicy.label', hint: 'actions.teams.groupPolicy.hint', notes: ['actions.teams.groupPolicy.note'], icon: <UsersRound size={16} /> },
  'teams.policies': { label: 'actions.teams.policies.label', hint: 'actions.teams.policies.hint', icon: <ListChecks size={16} /> },
  'teams.effectivePolicies': { label: 'actions.teams.effectivePolicies.label', hint: 'actions.teams.effectivePolicies.hint', icon: <BadgeCheck size={16} /> },
  'mail.purge': { label: 'actions.mail.purge.label', hint: 'actions.mail.purge.hint', notes: ['actions.mail.purge.note'], warn: ['actions.mail.purge.warn'], icon: <MailX size={16} /> },
  'mail.ruleAudit': { label: 'actions.mail.ruleAudit.label', hint: 'actions.mail.ruleAudit.hint', notes: ['actions.mail.ruleAudit.note'], icon: <Siren size={16} /> },
  'mail.reportThreat': { label: 'actions.mail.reportThreat.label', hint: 'actions.mail.reportThreat.hint', icon: <Flag size={16} /> },
  'mail.blockSender': { label: 'actions.mail.blockSender.label', hint: 'actions.mail.blockSender.hint', notes: ['actions.mail.blockSender.note'], icon: <ShieldBan size={16} /> },
  'mail.quarantine': { label: 'actions.mail.quarantine.label', hint: 'actions.mail.quarantine.hint', icon: <Archive size={16} /> },
  'mail.releaseQuarantine': { label: 'actions.mail.releaseQuarantine.label', hint: 'actions.mail.releaseQuarantine.hint', notes: ['actions.mail.releaseQuarantine.note'], icon: <ArchiveRestore size={16} /> },
  'group.membership': { label: 'groups.tileAdd', hint: 'groups.hintAdd', notes: ['groups.noteAdd'], icon: <UserPlus size={16} /> },
  'license.assign': { label: 'licensing.tileAssign', hint: 'licensing.hintAssign', notes: ['licensing.noteAssign'], warn: ['licensing.noteRemove'], icon: <Plus size={16} /> },
}

// Display tokens the backend puts into Change.before/after.
const VALUE_TOKENS = new Set(['allowed', 'blocked', 'active', 'revoked', 'shared', 'regular', 'enabled', 'disabled'])

// One fetch per connection (profile), shared by every page. A failed fetch is
// not cached: the next page that mounts asks again.
let cached: { key: string; list: Promise<CatalogEntry[]> } | null = null

// null until the first fetch for the current connection has landed.
export function useCatalog(): CatalogEntry[] | null {
  const { connected, status } = useStore()
  const key = connected ? `on:${status?.profileName ?? ''}` : 'off'
  const [list, setList] = useState<CatalogEntry[] | null>(null)
  useEffect(() => {
    if (cached?.key !== key) {
      const entry = { key, list: api.actions.catalog() }
      cached = entry
      entry.list.catch(() => { if (cached === entry) cached = null })
    }
    let alive = true
    setList(null)
    cached.list.then((v) => alive && setList(v), () => alive && setList([]))
    return () => { alive = false }
  }, [key])
  return list
}

// Expands catalog placeholders in a page's tile list into real tiles.
export function useCatalogTiles(pageId: string, actions: TaskAction[]) {
  const { t } = useTranslation()
  const loaded = useCatalog()
  const catalog = loaded ?? []
  const { status, mark } = useTaskStatus()
  const { askConfirm, confirmElement } = useConfirm()

  const tile = (e: CatalogEntry): TaskAction => {
    const ui = UI[e.id]
    const reason = e.reason ? t(`actions.reasons.${e.reason.key}`, { ...e.reason.params, defaultValue: e.reason.key }) : undefined
    const notes = [...(ui?.notes ?? []).map((k) => <p key={k}>{t(k)}</p>),
      ...(ui?.warn ?? []).map((k) => <p key={k} className="text-[var(--warn)]">{t(k)}</p>)]
    return {
      id: e.id,
      label: t(ui?.label ?? `actions.${e.id}.label`, { defaultValue: e.id }),
      hint: t(ui?.hint ?? `actions.${e.id}.hint`, { defaultValue: '' }) || undefined,
      icon: ui?.icon ?? <Zap size={16} />,
      variant: e.danger === 'destructive' ? 'danger' : undefined,
      write: e.danger !== 'read',
      badge: e.backend && e.backend !== 'graph' ? t(`actions.backends.${e.backend}`, { defaultValue: e.backend }) : undefined,
      disabledReason: e.available ? undefined : reason,
      warning: e.missingPermissions?.length ? t('actions.missingPermissions', { list: e.missingPermissions.join(', ') }) : undefined,
      note: notes.length ? <>{notes}</> : undefined,
      panel: e.available ? <CatalogPanel entry={e} mark={mark} askConfirm={askConfirm} /> : undefined,
    }
  }

  const byId = new Map(catalog.map((e) => [e.id, e]))
  const placed = new Set(actions.filter((a) => a.catalog).map((a) => a.id))
  const tiles = [
    ...actions.flatMap((a): TaskAction[] => {
      if (!a.catalog) return [a]
      const e = byId.get(a.id)
      if (e) return [tile(e)]
      // Still loading: hold the tile's place instead of making it jump in later.
      if (loaded === null && UI[a.id]) {
        return [{ id: a.id, label: t(UI[a.id].label), icon: UI[a.id].icon, disabledReason: t('common.loading') }]
      }
      return []
    }),
    ...catalog.filter((e) => e.page === pageId && !placed.has(e.id)).map(tile),
  ]
  return { tiles, status, confirmElement, ready: loaded !== null }
}

function CatalogPanel({ entry, mark, askConfirm }: {
  entry: CatalogEntry
  mark: (id: string, ok: boolean, text: string) => void
  askConfirm: (target: string, action: (confirm: string) => void) => void
}) {
  const { t } = useTranslation()
  const { readOnly, toast, cache, setCache } = useStore()
  // A read's last answer lives in the store cache with the values that
  // produced it, so leaving the page and coming back shows it again.
  const cacheKey = `catalog.read.${entry.id}`
  const cached = cache[cacheKey] as { values: Record<string, string>; result: engine.ReadResult } | undefined
  const initial = () => cached?.values ?? Object.fromEntries(entry.fields.map((f) => [f.name, f.default ?? '']))
  const [values, setValues] = useState<Record<string, string>>(initial)
  const [plan, setPlan] = useState<engine.Plan | null>(null)
  const [result, setResult] = useState<engine.Result | null>(null)
  const [rows, setRows] = useState<engine.ReadResult | null>(cached?.result ?? null)
  const [busy, setBusy] = useState(false)
  const isRead = entry.danger === 'read'
  // Only the newest read may show its rows: an edit or a second click makes
  // an answer still in flight stale.
  const runSeq = useRef(0)

  // Any edit invalidates the preview: Apply must run what is on screen.
  const set = (name: string, v: string) => {
    setValues((s) => ({ ...s, [name]: v }))
    setPlan(null)
    setResult(null)
    setRows(null)
    runSeq.current++
  }
  const ready = entry.fields.every((f) => !f.required || values[f.name])
  const actionable = !!plan?.changes?.some((c) => c.op !== 'none')
  const destructive = entry.danger === 'destructive'

  const run = async () => {
    const seq = ++runSeq.current
    setBusy(true)
    try {
      const r = await api.actions.run(entry.id, values)
      if (seq !== runSeq.current) return
      setRows(r)
      setCache(cacheKey, { values, result: r })
      mark(entry.id, true, t('actions.rows', { count: r.rows?.length ?? 0 }))
    } catch (e) {
      const m = errMessage(e)
      mark(entry.id, false, m)
      toast('err', m)
    } finally {
      setBusy(false)
    }
  }

  const preview = async () => {
    setBusy(true)
    setResult(null)
    try {
      setPlan(await api.actions.plan(entry.id, values))
    } catch (e) {
      toast('err', errMessage(e))
    } finally {
      setBusy(false)
    }
  }

  const apply = async (confirm: string) => {
    if (!plan) return
    setBusy(true)
    try {
      const r = await api.actions.apply(plan.id, confirm)
      setResult(r)
      setPlan(null)
      const ok = r.failed === 0
      const text = ok
        ? t('actions.result.ok', { applied: r.applied, skipped: r.skipped })
        : t('actions.result.failed', { failed: r.failed, total: r.outcomes.length })
      mark(entry.id, ok, text)
      toast(ok ? 'ok' : 'err', text)
    } catch (e) {
      // An expired preview cannot be retried as is — drop it so the button
      // reads "Preview" again.
      if (errParsed(e).code === 'planExpired') setPlan(null)
      const m = errMessage(e)
      mark(entry.id, false, m)
      toast('err', m)
    } finally {
      setBusy(false)
    }
  }

  const field = (f: engine.Field) => {
    const v = values[f.name] ?? ''
    const label = t(`actions.fields.${f.name}`, { defaultValue: f.name })
    switch (f.kind) {
      case 'user':
        return <EntityPicker value={v} onChange={(x) => set(f.name, x)} load={loadUsers} placeholder={t('users.pickUser')} />
      case 'group':
        return <EntityPicker value={v} onChange={(x) => set(f.name, x)} load={loadGroups} placeholder={t('groups.pickGroup')} />
      case 'sku':
        return <EntityPicker value={v} onChange={(x) => set(f.name, x)} load={loadSkus} placeholder={t('licensing.pickSku')} />
      case 'choice':
        return (
          <div role="radiogroup" aria-label={label} className="flex rounded-lg border border-[var(--border)] bg-[var(--bg)] p-0.5">
            {(f.options ?? []).map((o) => (
              <button key={o} type="button" role="radio" aria-checked={v === o} onClick={() => set(f.name, o)}
                className={`flex-1 rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                  v === o ? 'bg-[var(--accent)] text-[var(--accent-fg)]' : 'text-[var(--text-dim)] hover:text-[var(--text)]'}`}>
                {t(`actions.options.${o}`, { defaultValue: o })}
              </button>
            ))}
          </div>
        )
      default:
        return <Input value={v} onChange={(e) => set(f.name, e.target.value)} />
    }
  }

  return (
    <div className="flex flex-col gap-3">
      {entry.fields.map((f) => f.kind === 'choice'
        // Not a <label>: it would forward clicks on the caption to the first option.
        ? (
          <div key={f.name} className="flex flex-col gap-1">
            <span className="text-xs font-medium text-[var(--text-dim)]">{t(`actions.fields.${f.name}`, { defaultValue: f.name })}</span>
            {field(f)}
          </div>
        )
        : <Field key={f.name} label={t(`actions.fields.${f.name}`, { defaultValue: f.name })}>{field(f)}</Field>)}

      {isRead && (
        <Button variant="primary" disabled={busy || !ready} onClick={run}>
          {busy ? <Spinner /> : <Play size={15} />} {t('actions.run')}
        </Button>
      )}
      {isRead && rows && <RowsView result={rows} />}

      {!isRead && !plan && (
        <Button variant="primary" disabled={busy || !ready} onClick={preview}>
          {busy ? <Spinner /> : <Eye size={15} />} {t('actions.preview')}
        </Button>
      )}

      {plan && (
        <>
          <PlanView changes={plan.changes ?? []} />
          <div className="grid grid-cols-2 gap-2">
            <Button variant="subtle" disabled={busy} onClick={() => setPlan(null)}>{t('common.cancel')}</Button>
            <Button variant={destructive ? 'danger' : 'primary'} disabled={busy || readOnly || !actionable}
              onClick={() => (destructive ? askConfirm(plan.confirmTarget || '', apply) : apply(''))}>
              {busy ? <Spinner /> : <Check size={15} />} {t('actions.apply')}
            </Button>
          </div>
        </>
      )}

      {result && result.failed > 0 && (
        <ul className="flex flex-col gap-1 text-xs text-[var(--danger)]">
          {result.outcomes.filter((o) => !o.ok).map((o, i) => (
            <li key={i}>{o.target}: {o.error === 'canceled' ? t('common.canceled') : errMessage(o.error)}</li>
          ))}
        </ul>
      )}
    </div>
  )
}

// A read action's answer as a compact table. Column names and plain-word
// values (policy types, states) translate; names and addresses stay as is.
function RowsView({ result }: { result: engine.ReadResult }) {
  const { t, i18n } = useTranslation()
  const cols = result.columns ?? []
  const word = (v: string) => (/^[A-Za-z]+$/.test(v) && i18n.exists(`actions.values.${v}`) ? t(`actions.values.${v}`) : v)
  // "why" cells hold reason tokens: "externalForward=addr; deletes".
  const why = (v: string) => v.split('; ').map((tok) => {
    const i = tok.indexOf('=')
    const k = i < 0 ? tok : tok.slice(0, i)
    return t(`actions.why.${k}`, { value: i < 0 ? '' : tok.slice(i + 1), defaultValue: tok })
  }).join('; ')
  const show = (c: string, v: string) => (c === 'why' ? why(v) : word(v))
  const note = result.note && <p className="text-xs text-[var(--text-faint)]">{t(`actions.readNotes.${result.note.key}`, { ...result.note.params, defaultValue: result.note.key })}</p>
  if (!result.rows?.length) return <><p className="text-sm text-[var(--text-dim)]">{t('actions.noRows')}</p>{note}</>
  return (
    <>
    <div className="max-h-72 overflow-auto rounded-lg border border-[var(--border)] bg-[var(--bg)]">
      <table className="w-full text-left text-xs">
        <thead className="sticky top-0 bg-[var(--bg-elev)] text-[var(--text-faint)]">
          <tr>{cols.map((c) => <th key={c} className="px-2 py-1.5 font-medium">{t(`actions.columns.${c}`, { defaultValue: c })}</th>)}</tr>
        </thead>
        <tbody>
          {result.rows.map((r, i) => (
            <tr key={i} className="border-t border-[var(--border)]">
              {cols.map((c) => <td key={c} className="px-2 py-1.5">{show(c, r[c] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
    {note}
    </>
  )
}

// The preview: one row per change, "already so" rows dimmed.
function PlanView({ changes }: { changes: engine.Change[] }) {
  const { t } = useTranslation()
  const shown = (field: string, v?: string) => {
    if (!v) return '—'
    if (field === 'license') return skuFriendly(v)
    return VALUE_TOKENS.has(v) ? t(`actions.values.${v}`) : v
  }
  const nothing = changes.every((c) => c.op === 'none')
  return (
    <div className="rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3">
      <div className="mb-2 text-xs font-medium uppercase tracking-wide text-[var(--text-faint)]">{t('actions.planTitle')}</div>
      {nothing && <p className="mb-2 text-sm text-[var(--text-dim)]">{t('actions.nothingToDo')}</p>}
      <ul className="flex flex-col gap-1.5">
        {changes.map((c, i) => (
          <li key={i} className={`flex flex-wrap items-center gap-x-2 gap-y-0.5 text-sm ${c.op === 'none' ? 'opacity-50' : ''}`}>
            <span className={`flex items-center ${c.op === 'remove' ? 'text-[var(--danger)]' : c.op === 'none' ? 'text-[var(--text-faint)]' : 'text-[var(--ok)]'}`}>
              {c.op === 'remove' ? <Minus size={13} /> : c.op === 'add' ? <Plus size={13} /> : <ArrowRight size={13} />}
            </span>
            <span className="font-medium">{c.target}</span>
            <span className="text-[var(--text-faint)]">{t(`actions.changeFields.${c.field}`, { defaultValue: c.field })}</span>
            <span className="text-[var(--text-dim)]">
              {c.op === 'add' ? shown(c.field, c.after)
                : c.op === 'remove' ? <s>{shown(c.field, c.before)}</s>
                : c.op === 'none' ? `${shown(c.field, c.after || c.before)} · ${t('actions.op.none')}`
                : <>{shown(c.field, c.before)} → <span className="text-[var(--text)]">{shown(c.field, c.after)}</span></>}
            </span>
            {c.note && <span className="w-full text-xs text-[var(--warn)]">{t(`actions.notes.${c.note}`, { defaultValue: c.note })}</span>}
          </li>
        ))}
      </ul>
    </div>
  )
}
