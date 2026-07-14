import type { CSSProperties } from 'react'

const DESIRED = 320

// Positions a portal dropdown directly below its trigger. Height is capped to
// the space available below so it scrolls internally instead of leaving the
// window (a webview can't paint outside its bounds).
export function dropdownStyle(rect: DOMRect, opts?: { minWidth?: number; maxWidth?: number }): CSSProperties {
  const margin = 8
  const spaceBelow = window.innerHeight - rect.bottom - margin
  const maxHeight = Math.max(160, Math.min(DESIRED, spaceBelow))

  let width = rect.width
  if (opts?.minWidth) width = Math.max(width, opts.minWidth)
  if (opts?.maxWidth) width = Math.min(width, opts.maxWidth)

  return { position: 'fixed', top: rect.bottom + 4, left: rect.left, width, maxHeight }
}

function scrollParent(el: HTMLElement): HTMLElement | null {
  let node = el.parentElement
  while (node) {
    const oy = getComputedStyle(node).overflowY
    if ((oy === 'auto' || oy === 'scroll') && node.scrollHeight > node.clientHeight) return node
    node = node.parentElement
  }
  return null
}

// When the trigger is near the bottom, scroll its container up so a downward
// dropdown of ~DESIRED px fits fully in view.
export function nudgeIntoView(el: HTMLElement) {
  const rect = el.getBoundingClientRect()
  const spaceBelow = window.innerHeight - rect.bottom - 8
  const need = DESIRED - spaceBelow
  if (need <= 0) return
  const sp = scrollParent(el)
  if (sp) sp.scrollBy({ top: need + 12 })
  else window.scrollBy({ top: need + 12 })
}
