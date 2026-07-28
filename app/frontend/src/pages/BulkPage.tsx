import { useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Play, FileUp, FileDown, CheckCircle2, XCircle, UserPlus, KeyRound, Users } from 'lucide-react'
import { TaskPage, TaskForm, type TaskAction } from '../components/TaskPage'
import { Button, Field, Spinner, Badge } from '../components/ui'
import { JobConsole } from '../components/JobConsole'
import { useConfirm } from '../lib/useConfirm'
import { useStore, type BulkItem, type BulkRowResult } from '../lib/store'
import { api, errMessage, type GraphObject } from '../lib/api'
import { skuFriendly } from '../lib/skuNames'
import { parseCsv } from '../lib/csv'

type OpId = 'createUsers' | 'assignLicense' | 'addToGroup'

// Nobody keeps SKU or group GUIDs at hand, so these columns accept a product /
// group NAME as well and resolve it once per run. A GUID still passes through.
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i

async function skuResolver(): Promise<(v: string) => string> {
  const skus = await api.licensing.skus()
  const byName = new Map<string, string>()
  for (const s of skus as GraphObject[]) {
    // A SKU without a part number would blow up skuFriendly and take the whole
    // run with it; skip its name, keep it addressable by GUID.
    if (!s.skuPartNumber) continue
    byName.set(String(s.skuPartNumber).toLowerCase(), s.skuId)
    byName.set(skuFriendly(s.skuPartNumber).toLowerCase(), s.skuId)
  }
  return (v: string) => {
    if (GUID.test(v)) return v
    const hit = byName.get(v.trim().toLowerCase())
    if (!hit) throw new Error(`unknown license: ${v}`)
    return hit
  }
}

async function groupResolver(): Promise<(v: string) => string> {
  const groups = await api.groups.list('', 0)
  const byName = new Map<string, string>()
  for (const g of groups as GraphObject[]) {
    if (g.displayName) byName.set(String(g.displayName).toLowerCase(), g.id)
    if (g.mail) byName.set(String(g.mail).toLowerCase(), g.id)
  }
  return (v: string) => {
    if (GUID.test(v)) return v
    const hit = byName.get(v.trim().toLowerCase())
    if (!hit) throw new Error(`unknown group: ${v}`)
    return hit
  }
}

// Each operation: CSV columns (order-independent, matched by header name), which
// of them are required, an optional name→id resolver, and the call for one row.
const OPS: Record<OpId, {
  headers: string[]
  required: string[]
  resolve?: () => Promise<(v: string) => string>
  run: (r: Record<string, string>, resolve?: (v: string) => string) => Promise<string | void>
}> = {
  createUsers: {
    headers: ['displayName', 'upn', 'mailNickname', 'password', 'usageLocation'],
    required: ['displayName', 'upn', 'mailNickname', 'password'],
    run: async (r) => {
      await api.users.create(r.displayName, r.upn, r.mailNickname, r.password, true, r.usageLocation || '')
    },
  },
  assignLicense: {
    headers: ['upn', 'license'],
    required: ['upn', 'license'],
    resolve: skuResolver,
    run: async (r, resolve) => {
      await api.licensing.assign(r.upn, [resolve!(r.license)], [])
    },
  },
  addToGroup: {
    headers: ['upn', 'group'],
    required: ['upn', 'group'],
    resolve: groupResolver,
    run: async (r, resolve) => {
      await api.groups.addMember(resolve!(r.group), r.upn)
    },
  },
}

export function BulkPage() {
  const { t } = useTranslation()
  const { readOnly, toast, jobs, startBulkRun, cancelBulkRun, clearJob } = useStore()
  const { askConfirm, confirmElement } = useConfirm()
  const fileRef = useRef<HTMLInputElement>(null)

  const [op, setOp] = useState<OpId>('createUsers')
  const [csv, setCsv] = useState('')
  const [resolving, setResolving] = useState(false)

  const job = jobs.bulk
  const busy = !!job?.running
  const results = (job?.result as BulkRowResult[] | null) || null

  const spec = OPS[op]

  // Parse + validate the CSV against the selected operation's headers.
  const parsed = (() => {
    if (!csv.trim()) return { rows: [] as Record<string, string>[], errors: [] as string[] }
    const raw = parseCsv(csv)
    if (raw.length < 2) return { rows: [], errors: raw.length ? [t('bulk.needData')] : [] }
    const header = raw[0].map((h) => h.trim())
    const unknown = header.filter((h) => !spec.headers.includes(h))
    const missing = spec.required.filter((h) => !header.includes(h))
    const errors: string[] = []
    if (unknown.length) errors.push(t('bulk.unknownColumns', { cols: unknown.join(', ') }))
    if (missing.length) errors.push(t('bulk.missingColumns', { cols: missing.join(', ') }))
    if (errors.length) return { rows: [], errors }
    const rows = raw.slice(1).map((cells) => {
      const rec: Record<string, string> = {}
      header.forEach((h, i) => { rec[h] = (cells[i] ?? '').trim() })
      return rec
    })
    rows.forEach((r, i) => {
      const bad = spec.required.filter((h) => !r[h])
      if (bad.length) errors.push(t('bulk.rowMissing', { n: i + 2, cols: bad.join(', ') }))
    })
    return { rows: errors.length ? [] : rows, errors }
  })()

  const openFile = (f: File | undefined) => {
    if (!f) return
    f.text().then(setCsv).catch(() => toast('err', 'read failed'))
  }

  const downloadTemplate = () => {
    const blob = new Blob([spec.headers.join(',') + '\n'], { type: 'text/csv' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `swissknife-${op}.csv`
    a.click()
    URL.revokeObjectURL(a.href)
  }

  const run = () => {
    askConfirm('RUN', async () => {
      // Names are resolved once, before the run, so a typo fails immediately
      // instead of after twenty half-applied rows.
      let resolve: ((v: string) => string) | undefined
      if (spec.resolve) {
        setResolving(true)
        try { resolve = await spec.resolve() } catch (e) { toast('err', errMessage(e)); return }
        finally { setResolving(false) }
      }
      const items: BulkItem[] = parsed.rows.map((r) => ({
        label: r.upn || r.displayName || JSON.stringify(r),
        run: () => spec.run(r, resolve),
      }))
      startBulkRun(items)
    }, t('bulk.confirm', { n: parsed.rows.length, op: t(`bulk.op.${op}`) }))
  }

  // One tile per operation: the CSV form is the same, only the columns differ.
  const opPanel = (id: OpId) => (
    <TaskForm>
      <div className="flex flex-wrap gap-2">
        <Button onClick={downloadTemplate}><FileDown size={15} /> {t('bulk.template')}</Button>
        <Button onClick={() => fileRef.current?.click()}><FileUp size={15} /> {t('bulk.openCsv')}</Button>
        <input ref={fileRef} type="file" accept=".csv,text/csv" className="hidden"
          onChange={(e) => { openFile(e.target.files?.[0]); e.target.value = '' }} />
      </div>
      <Field label={t('bulk.csv')} hint={OPS[id].headers.join(', ')}>
        <textarea
          value={csv} onChange={(e) => setCsv(e.target.value)}
          placeholder={OPS[id].headers.join(',') + '\n…'}
          spellCheck={false}
          className="h-40 w-full resize-y rounded-lg border border-[var(--border)] bg-[var(--bg)] p-3 font-mono text-xs text-[var(--text)] outline-none focus:border-[var(--accent)]" />
      </Field>
      {parsed.errors.length > 0 && (
        <div className="flex flex-col gap-1">
          {parsed.errors.slice(0, 6).map((e, i) => <div key={i} className="text-xs text-[var(--danger)]">{e}</div>)}
        </div>
      )}
      {parsed.rows.length > 0 && <div className="text-xs text-[var(--text-dim)]">{t('bulk.parsed', { n: parsed.rows.length })}</div>}
      <Button variant="primary" disabled={readOnly || busy || resolving || parsed.rows.length === 0} onClick={run}>
        {busy || resolving ? <Spinner /> : <Play size={15} />}
        {resolving ? t('bulk.resolving') : t('bulk.run', { n: parsed.rows.length })}
      </Button>
    </TaskForm>
  )

  const actions: TaskAction[] = (Object.keys(OPS) as OpId[]).map((id) => ({
    id,
    label: t(`bulk.tile.${id}`),
    hint: t(`bulk.hint.${id}`),
    icon: id === 'createUsers' ? <UserPlus size={16} /> : id === 'assignLicense' ? <KeyRound size={16} /> : <Users size={16} />,
    variant: 'primary' as const,
    write: true,
    note: <p>{t(`bulk.note.${id}`)}</p>,
    // Selecting a tile switches which columns the CSV is validated against.
    onClick: () => setOp(id),
    panel: opPanel(id),
  }))

  const resultPane = (
    <div className="flex h-full flex-col gap-3 overflow-auto p-3">
      {job && (job.running || job.log.length > 0) && (
        <JobConsole job={job} onCancel={cancelBulkRun} onClear={() => clearJob('bulk')} />
      )}
      {results && results.length > 0 && (
        <>
          <div className="flex gap-2">
            <Badge kind="ok">{t('bulk.ok')}: {results.filter((r) => r.ok).length}</Badge>
            <Badge kind="danger">{t('bulk.failed')}: {results.filter((r) => !r.ok).length}</Badge>
          </div>
          <div className="flex flex-col gap-1">
            {results.map((r, i) => (
              <div key={i} className="flex items-start gap-2 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-1.5 text-sm">
                {r.ok ? <CheckCircle2 size={15} className="mt-0.5 shrink-0 text-[var(--ok)]" /> : <XCircle size={15} className="mt-0.5 shrink-0 text-[var(--danger)]" />}
                <div className="min-w-0">
                  <span className="font-medium">{r.label}</span>
                  {r.error && <span className="ml-2 text-xs text-[var(--danger)]">{r.error}</span>}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  )

  return (
    <>
      {confirmElement}
      <TaskPage
        pageId="bulk"
        title={t('bulk.title')}
        subtitle={t('bulk.subtitle')}
        actions={actions}
        busy={busy || resolving}
        busyLabel={resolving ? t('bulk.resolving') : job?.progress}
        hasResult={!!job && (job.running || job.log.length > 0 || !!results)}
        onClearResult={() => clearJob('bulk')}
        result={resultPane}
      />
    </>
  )
}
