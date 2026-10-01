import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { Search } from 'lucide-react'
import { Button, Field, Input, Spinner, ErrorNote } from './ui'
import { UpnInput } from './UpnInput'
import { useStore } from '../lib/store'
import { api, errMessage } from '../lib/api'
import type { services } from '../../wailsjs/go/models'

// The mailbox copy options, shared by the offboarding playbook form and the
// standalone tile on the Offboarding page so both ask the same questions.
export type MailboxCopyOptions = {
  target: string
  folder: string
  includeContacts: boolean
  includeCalendar: boolean
}

export function MailboxCopyFields({
  source,
  value,
  onChange,
  disabled,
}: {
  source: string // the mailbox owner — the preview needs it
  value: MailboxCopyOptions
  onChange: (v: MailboxCopyOptions) => void
  disabled?: boolean
}) {
  const { t } = useTranslation()
  const { cache, setCache } = useStore()
  // A preview walks the whole folder tree; keeping it in the store means
  // leaving the page and coming back does not throw it away.
  const cacheKey = `mailbox.preview.${source}`
  const preview = (cache[cacheKey] as services.MailboxPreview | undefined) ?? null
  const [previewing, setPreviewing] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const runPreview = async () => {
    setPreviewing(true); setError(null)
    try {
      setCache(cacheKey, await api.mailboxTransfer.preview(source))
    } catch (e) {
      setError(errMessage(e))
    } finally { setPreviewing(false) }
  }

  const set = (patch: Partial<MailboxCopyOptions>) => onChange({ ...value, ...patch })
  const folders = preview
    ? preview.mailFolders + (value.includeContacts ? preview.contactFolders : 0) + (value.includeCalendar ? preview.calendarFolders : 0)
    : 0
  const items = preview
    ? preview.mailItems + (value.includeContacts ? preview.contactItems : 0) + (value.includeCalendar ? preview.calendarItems : 0)
    : 0

  return (
    <>
      <Field label={t('mailboxTransfer.copyTo')}>
        <UpnInput value={value.target} onChange={(v) => set({ target: v })} placeholder="manager@contoso.com" />
      </Field>
      <Field label={t('mailboxTransfer.folder')} hint={t('mailboxTransfer.folderHint')}>
        <Input value={value.folder} onChange={(e) => set({ folder: e.target.value })} disabled={disabled} />
      </Field>
      <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
        <input type="checkbox" checked={value.includeContacts} disabled={disabled} onChange={(e) => set({ includeContacts: e.target.checked })} />
        {t('mailboxTransfer.includeContacts')}
      </label>
      <label className="flex items-center gap-2 text-sm text-[var(--text-dim)]">
        <input type="checkbox" checked={value.includeCalendar} disabled={disabled} onChange={(e) => set({ includeCalendar: e.target.checked })} />
        {t('mailboxTransfer.includeCalendar')}
      </label>
      <div className="flex flex-wrap items-center gap-2">
        <Button variant="subtle" disabled={!source || previewing} onClick={runPreview}>
          {previewing ? <Spinner /> : <Search size={14} />} {t('mailboxTransfer.preview')}
        </Button>
        {!source && <span className="text-xs text-[var(--text-faint)]">{t('mailboxTransfer.needsSource')}</span>}
      </div>
      {preview && (
        <div className="text-xs text-[var(--text-dim)]">
          <p>{t('mailboxTransfer.previewResult', {
            folders, items, mail: preview.mailItems, contacts: preview.contactItems, calendar: preview.calendarItems,
          })}</p>
          {(preview.systemFolders > 0 || preview.otherItems > 0) && (
            <p className="text-[var(--text-faint)]">
              {t('mailboxTransfer.previewSkipped', { system: preview.systemFolders, other: preview.otherItems })}
            </p>
          )}
        </div>
      )}
      {error && <ErrorNote>{error}</ErrorNote>}
    </>
  )
}

// What the operator must know before running: which API this is, what it
// is not (a backup product), and which permissions it takes.
export function MailboxCopyNote() {
  const { t } = useTranslation()
  return (
    <>
      <p>{t('mailboxTransfer.note')}</p>
      <p className="text-[var(--warn)]">{t('mailboxTransfer.permissionNote')}</p>
    </>
  )
}
