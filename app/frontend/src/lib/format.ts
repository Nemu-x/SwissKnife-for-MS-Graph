// Flattening of Graph objects for tables + masking (safe mode).

const SENSITIVE = ['token', 'secret', 'password', 'authorization', 'access', 'refresh']

export function maskSensitive(obj: any): any {
  if (Array.isArray(obj)) return obj.map(maskSensitive)
  if (obj && typeof obj === 'object') {
    const out: any = {}
    for (const [k, v] of Object.entries(obj)) {
      out[k] = SENSITIVE.some((s) => k.toLowerCase().includes(s)) ? '***MASKED***' : maskSensitive(v)
    }
    return out
  }
  return obj
}

// Picks the interesting columns for a table from an array of objects.
export function pickColumns(rows: GraphRow[], limit = 8): string[] {
  const priority = [
    'displayName', 'name', 'userPrincipalName', 'mail', 'subject', 'topic',
    'id', 'accountEnabled', 'deviceName', 'operatingSystem', 'createdDateTime',
    'receivedDateTime', 'skuPartNumber',
  ]
  const seen = new Set<string>()
  for (const r of rows) for (const k of Object.keys(r)) if (!k.startsWith('@')) seen.add(k)
  const cols = [...priority.filter((p) => seen.has(p))]
  for (const k of seen) if (!cols.includes(k) && cols.length < limit) cols.push(k)
  return cols.slice(0, limit)
}

export type GraphRow = Record<string, any>

export function cellText(v: any): string {
  if (v == null) return ''
  if (typeof v === 'object') {
    if ('displayName' in v) return String(v.displayName)
    if ('emailAddress' in v && v.emailAddress?.address) return String(v.emailAddress.address)
    return JSON.stringify(v)
  }
  return String(v)
}

export function toCSV(cols: string[], rows: GraphRow[]): string {
  const esc = (s: string) => (/[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s)
  const head = cols.map(esc).join(',')
  const body = rows.map((r) => cols.map((c) => esc(cellText(r[c]))).join(',')).join('\n')
  return '﻿' + head + '\n' + body
}

export function humanBytes(n: number): string {
  if (!n || n < 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  let i = 0
  let v = n
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024
    i++
  }
  return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`
}

export function downloadText(filename: string, text: string, mime = 'text/csv') {
  const blob = new Blob([text], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}
