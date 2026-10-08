import { useRef, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { X } from 'lucide-react'
import { Button, Field, Select } from './ui'
import { EntityPicker } from './EntityPicker'
import { useStore } from '../lib/store'
import { api, errMessage, type Profile } from '../lib/api'
import { loadGroups } from '../lib/pickers'
import type { Option } from './MultiSelect'

type Policy = { maxDanger?: string; allowedGroups?: string[]; groupLabels?: Record<string, string> }

// Per-profile guard rails: the highest danger level and the groups whose
// members the profile may change. Enforced by the backend once connected.
export function ProfileLimits({ profile, onSaved }: { profile: Profile; onSaved: () => void }) {
  const { t } = useTranslation()
  const { connected, toast, refreshStatus } = useStore()
  const init = ((profile as any).policy || {}) as Policy
  const [maxDanger, setMaxDanger] = useState(init.maxDanger || '')
  const [groups, setGroups] = useState<string[]>(init.allowedGroups || [])
  const [labels, setLabels] = useState<Record<string, string>>(init.groupLabels || {})
  const opts = useRef<Option[]>([])
  const [busy, setBusy] = useState(false)

  const add = (id: string) => {
    if (!id || groups.includes(id)) return
    const o = opts.current.find((x) => x.value === id)
    setGroups([...groups, id])
    if (o) setLabels({ ...labels, [id]: o.label })
  }
  const save = async () => {
    setBusy(true)
    try {
      const kept: Record<string, string> = {}
      groups.forEach((g) => { if (labels[g]) kept[g] = labels[g] })
      await api.connect.setProfilePolicy(profile.id, { maxDanger, allowedGroups: groups, groupLabels: kept })
      toast('ok', t('common.save'))
      onSaved()
      refreshStatus()
    } catch (e) { toast('err', errMessage(e)) } finally { setBusy(false) }
  }

  return (
    <div className="flex flex-col gap-2 rounded-lg border border-[var(--border)] bg-[var(--surface)] px-3 py-2">
      <p className="text-xs text-[var(--text-faint)]">{t('connect.limits.hint')}</p>
      <Field label={t('connect.limits.maxDanger')}>
        <Select value={maxDanger} onChange={(e) => setMaxDanger(e.target.value)} className="w-full">
          <option value="">{t('connect.limits.noLimit')}</option>
          <option value="read">{t('connect.limits.read')}</option>
          <option value="write">{t('connect.limits.write')}</option>
          <option value="destructive">{t('connect.limits.destructive')}</option>
        </Select>
      </Field>
      <Field label={t('connect.limits.groups')}>
        <div className="flex flex-wrap gap-1">
          {groups.map((g) => (
            <span key={g} className="inline-flex items-center gap-1 rounded-md bg-[var(--bg)] px-2 py-0.5 text-xs">
              {labels[g] || g}
              <button onClick={() => setGroups(groups.filter((x) => x !== g))} aria-label={t('common.remove')}><X size={11} /></button>
            </span>
          ))}
          {groups.length === 0 && <span className="text-xs text-[var(--text-faint)]">{t('connect.limits.anyTarget')}</span>}
        </div>
        {connected
          ? <EntityPicker value="" onChange={add} placeholder={t('connect.limits.addGroup')}
              load={async () => { const o = await loadGroups(); opts.current = o; return o }} />
          : <p className="text-xs text-[var(--text-faint)]">{t('connect.limits.connectToPick')}</p>}
      </Field>
      <Button variant="primary" onClick={save} disabled={busy}>{t('common.save')}</Button>
    </div>
  )
}
