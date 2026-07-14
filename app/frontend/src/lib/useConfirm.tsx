import { useState, useCallback } from 'react'
import { ConfirmDestructive } from '../components/ConfirmDestructive'

// Hook for the destructive flow: askConfirm(target, action) opens the modal;
// after a correct target is entered, it calls action(confirmText).
export function useConfirm() {
  const [state, setState] = useState<{
    target: string
    title?: string
    action: (confirm: string) => void
  } | null>(null)

  const askConfirm = useCallback((target: string, action: (confirm: string) => void, title?: string) => {
    setState({ target, action, title })
  }, [])

  const element = state ? (
    <ConfirmDestructive
      target={state.target}
      title={state.title}
      onConfirm={(c) => {
        const a = state.action
        setState(null)
        a(c)
      }}
      onCancel={() => setState(null)}
    />
  ) : null

  return { askConfirm, confirmElement: element }
}
