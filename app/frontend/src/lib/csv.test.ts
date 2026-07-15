import { describe, it, expect } from 'vitest'
import { parseCsv } from './csv'

describe('parseCsv', () => {
  it('parses plain rows', () => {
    expect(parseCsv('a,b\n1,2')).toEqual([['a', 'b'], ['1', '2']])
  })

  it('handles CRLF and trailing newline', () => {
    expect(parseCsv('a,b\r\n1,2\r\n')).toEqual([['a', 'b'], ['1', '2']])
  })

  it('handles quoted fields with commas and escaped quotes', () => {
    expect(parseCsv('name,note\n"Doe, John","said ""hi"""')).toEqual([
      ['name', 'note'],
      ['Doe, John', 'said "hi"'],
    ])
  })

  it('keeps empty cells but drops blank lines', () => {
    expect(parseCsv('a,b\n\n1,\n,,\n')).toEqual([['a', 'b'], ['1', '']])
  })

  it('handles newlines inside quotes', () => {
    expect(parseCsv('a\n"line1\nline2"')).toEqual([['a'], ['line1\nline2']])
  })

  it('returns nothing for empty input', () => {
    expect(parseCsv('')).toEqual([])
    expect(parseCsv('\n\n')).toEqual([])
  })
})
