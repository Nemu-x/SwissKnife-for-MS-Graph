import { useState } from 'react'
import { useTranslation } from 'react-i18next'
import { AlertTriangle } from 'lucide-react'
import { Button, Input } from './ui'

// Typed-confirm modal: the user must type the target (the backend re-verifies).
export function ConfirmDestructive({
  target,
  title,
  onConfirm,
  onCancel,
}: {
  target: string
  title?: string
  onConfirm: (confirm: string) => void
  onCancel: () => void
}) {
  const { t } = useTranslation()
  const [text, setText] = useState('')
  const matches = text === target

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4" onClick={onCancel}>
      <div
        className="w-full max-w-md rounded-xl border border-[var(--danger)]/40 bg-[var(--bg-elev)] p-5 shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex items-center gap-2 text-[var(--danger)]">
          <AlertTriangle size={20} />
          <h3 className="text-base font-semibold">{title || t('safety.destructiveTitle')}</h3>
        </div>
        <p className="mb-3 text-sm text-[var(--text-dim)]">
          {t('safety.destructiveBody', { target })}
        </p>
        <Input
          autoFocus
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder={t('safety.typeToConfirm')}
        />
        <div className="mt-4 flex justify-end gap-2">
          <Button variant="ghost" onClick={onCancel}>
            {t('common.cancel')}
          </Button>
          <Button variant="danger" disabled={!matches} onClick={() => onConfirm(text)}>
            {t('common.confirm')}
          </Button>
        </div>
      </div>
    </div>
  )
}
