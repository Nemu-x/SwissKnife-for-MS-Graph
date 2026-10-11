import { describe, expect, it, vi } from 'vitest'

vi.mock('../../wailsjs/runtime/runtime', () => ({ BrowserOpenURL: vi.fn() }))

import { safeHttpsUrl } from './links'

describe('safeHttpsUrl', () => {
  it('accepts https links', () => {
    expect(safeHttpsUrl('https://entra.microsoft.com/#view/x')).toBe('https://entra.microsoft.com/#view/x')
  })
  it('rejects every other scheme and junk', () => {
    for (const bad of ['javascript:alert(1)', 'http://example.com', 'file:///C:/Windows/System32/calc.exe',
      'ms-settings:', 'data:text/html,x', '/relative', '', '   ', null, undefined, 42, 'https://u:p@example.com']) {
      expect(safeHttpsUrl(bad)).toBeNull()
    }
  })
})
