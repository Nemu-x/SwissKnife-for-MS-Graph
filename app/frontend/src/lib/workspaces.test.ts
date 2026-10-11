import { describe, expect, it } from 'vitest'
import { homePage, pageEnabled } from './workspaces'

describe('workspaces', () => {
  it('shows only the parts that are switched on', () => {
    const cloud = { cloud: true, onprem: false }
    const onprem = { cloud: false, onprem: true }
    const both = { cloud: true, onprem: true }
    expect(pageEnabled('onprem', cloud)).toBe(false)
    expect(pageEnabled('users', cloud)).toBe(true)
    expect(pageEnabled('users', onprem)).toBe(false)
    expect(pageEnabled('connect', onprem)).toBe(false)
    expect(pageEnabled('onprem', both) && pageEnabled('users', both)).toBe(true)
    // Settings and run history belong to everyone.
    expect(pageEnabled('settings', onprem) && pageEnabled('history', cloud)).toBe(true)
    expect(homePage(onprem)).toBe('onprem')
    expect(homePage(both)).toBe('connect')
  })
})
