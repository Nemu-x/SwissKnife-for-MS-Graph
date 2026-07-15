import { describe, it, expect } from 'vitest'
import en from './en'
import ru from './ru'

// Every UI string must exist in BOTH locales — a key present in one and not
// the other means an untranslated (or silently missing) string in the app.
function keysOf(obj: Record<string, unknown>, prefix = ''): string[] {
  return Object.entries(obj).flatMap(([k, v]) =>
    v !== null && typeof v === 'object'
      ? keysOf(v as Record<string, unknown>, `${prefix}${k}.`)
      : [`${prefix}${k}`],
  )
}

describe('locale parity (en ↔ ru)', () => {
  const enKeys = keysOf(en).sort()
  const ruKeys = keysOf(ru as Record<string, unknown>).sort()

  it('ru has every en key', () => {
    const missing = enKeys.filter((k) => !ruKeys.includes(k))
    expect(missing).toEqual([])
  })

  it('en has every ru key', () => {
    const extra = ruKeys.filter((k) => !enKeys.includes(k))
    expect(extra).toEqual([])
  })

  it('interpolation placeholders match', () => {
    const flat = (o: Record<string, unknown>, p = ''): Record<string, string> =>
      Object.entries(o).reduce((acc, [k, v]) => {
        if (v !== null && typeof v === 'object') Object.assign(acc, flat(v as Record<string, unknown>, `${p}${k}.`))
        else acc[`${p}${k}`] = String(v)
        return acc
      }, {} as Record<string, string>)
    const enFlat = flat(en)
    const ruFlat = flat(ru as Record<string, unknown>)
    const holes = (s: string) => [...s.matchAll(/\{\{(\w+)\}\}/g)].map((m) => m[1]).sort()
    const mismatched = Object.keys(enFlat)
      .filter((k) => k in ruFlat)
      .filter((k) => JSON.stringify(holes(enFlat[k])) !== JSON.stringify(holes(ruFlat[k])))
    expect(mismatched).toEqual([])
  })
})
