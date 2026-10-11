import { BrowserOpenURL } from '../../wailsjs/runtime/runtime'

// safeHttpsUrl returns the normalized URL when it is an absolute https: link,
// otherwise null. URLs that come from Graph responses are untrusted: a
// javascript:, file: or custom-scheme link must never be rendered clickable.
export function safeHttpsUrl(raw: unknown): string | null {
  if (typeof raw !== 'string' || !raw.trim()) return null
  try {
    const u = new URL(raw.trim())
    if (u.protocol !== 'https:' || u.username || u.password) return null
    return u.toString()
  } catch {
    return null
  }
}

// openExternal opens a vetted https URL in the system browser (never inside
// the webview).
export function openExternal(raw: unknown) {
  const u = safeHttpsUrl(raw)
  if (u) BrowserOpenURL(u)
}
