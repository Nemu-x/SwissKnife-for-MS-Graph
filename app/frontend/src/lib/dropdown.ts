import type { CSSProperties } from 'react'

// Positions a portal dropdown next to its trigger: opens downward, but flips
// upward when there isn't enough room below. Height is capped to the available
// space so it scrolls internally and never leaves the window.
export function dropdownStyle(rect: DOMRect, opts?: { minWidth?: number; maxWidth?: number }): CSSProperties {
  const margin = 8
  const spaceBelow = window.innerHeight - rect.bottom - margin
  const spaceAbove = rect.top - margin
  const openUp = spaceBelow < 260 && spaceAbove > spaceBelow
  const maxHeight = Math.max(160, Math.min(340, openUp ? spaceAbove : spaceBelow))

  let width = rect.width
  if (opts?.minWidth) width = Math.max(width, opts.minWidth)
  if (opts?.maxWidth) width = Math.min(width, opts.maxWidth)

  const style: CSSProperties = { position: 'fixed', left: rect.left, width, maxHeight }
  if (openUp) style.bottom = window.innerHeight - rect.top + 4
  else style.top = rect.bottom + 4
  return style
}
