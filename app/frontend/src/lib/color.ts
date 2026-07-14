// Small color helpers to derive accent variants from a single hex.

export interface AccentPreset {
  name: string
  hex: string
}

export const ACCENT_PRESETS: AccentPreset[] = [
  { name: 'Indigo', hex: '#6366f1' },
  { name: 'Violet', hex: '#8b5cf6' },
  { name: 'Sky', hex: '#0ea5e9' },
  { name: 'Teal', hex: '#14b8a6' },
  { name: 'Emerald', hex: '#10b981' },
  { name: 'Amber', hex: '#f59e0b' },
  { name: 'Rose', hex: '#f43f5e' },
  { name: 'Slate', hex: '#64748b' },
]

export const DEFAULT_ACCENT = ACCENT_PRESETS[0].hex

function clamp(n: number) {
  return Math.max(0, Math.min(255, Math.round(n)))
}

export function isHex(v: string): boolean {
  return /^#([0-9a-fA-F]{6})$/.test(v.trim())
}

function toRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}

function toHex(r: number, g: number, b: number): string {
  return '#' + [r, g, b].map((n) => clamp(n).toString(16).padStart(2, '0')).join('')
}

// Mix a color toward black (amount<0) or white (amount>0), amount in [-1, 1].
function shade(hex: string, amount: number): string {
  const [r, g, b] = toRgb(hex)
  const t = amount < 0 ? 0 : 255
  const p = Math.abs(amount)
  return toHex(r + (t - r) * p, g + (t - g) * p, b + (t - b) * p)
}

// Relative luminance to pick readable foreground.
function luminance(hex: string): number {
  const [r, g, b] = toRgb(hex).map((c) => {
    const s = c / 255
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4
  })
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}

const ACCENT_VARS = ['--accent', '--accent-hover', '--accent2', '--ring', '--accent-fg']

// Removes inline overrides so the theme's palette accent applies ("Auto").
export function clearAccent() {
  const root = document.documentElement
  ACCENT_VARS.forEach((v) => root.style.removeProperty(v))
}

// Applies a custom accent hex on top of the theme. Empty/invalid => Auto.
export function applyAccent(hex: string) {
  if (!isHex(hex)) {
    clearAccent()
    return
  }
  const root = document.documentElement
  root.style.setProperty('--accent', hex)
  root.style.setProperty('--accent-hover', shade(hex, -0.15))
  root.style.setProperty('--accent2', shade(hex, 0.2))
  root.style.setProperty('--ring', hex)
  root.style.setProperty('--accent-fg', luminance(hex) > 0.55 ? '#0b0d12' : '#ffffff')
}
