import { describe, it, expect } from 'vitest'
import { toCSV } from './format'

describe('toCSV', () => {
  it('keeps formula-looking cells as text', () => {
    const csv = toCSV(['name'], [{ name: '=HYPERLINK("http://evil")' }, { name: '+1' }, { name: '-2' }, { name: '@x' }, { name: 'Ann' }])
    const lines = csv.replace(/^\uFEFF/, '').split('\n')
    expect(lines.slice(1)).toEqual([`"'=HYPERLINK(""http://evil"")"`, "'+1", "'-2", "'@x", 'Ann'])
  })

  it('quotes a cell with a carriage return so it cannot split the record', () => {
    const csv = toCSV(['name'], [{ name: 'a\r=1+1' }])
    expect(csv.replace(/^﻿/, '').split('\n')[1]).toBe('"a\r=1+1"')
  })
})
