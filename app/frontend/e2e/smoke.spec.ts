import { test, expect, type Page } from '@playwright/test'

// The app normally runs inside Wails, which injects window.go (bindings) and
// window.runtime (events). We stub both so the real production bundle can be
// exercised in a plain browser: any binding call resolves with a sane default.
// This catches white-screen crashes, i18n init breakage, and nav regressions.
async function stubWails(page: Page, opts: { connected?: boolean } = {}) {
  await page.addInitScript((connected: boolean) => {
    const results: Record<string, unknown> = {
      GetStatus: connected ? { connected: true, profileName: 'e2e', readOnly: false } : null,
      Profiles: [],
      Domains: [],
      Version: '0.0.0-e2e',
      Probe: {},
      Check: { currentVersion: '0.0.0-e2e', latestVersion: '', updateAvailable: false, notes: '', url: '' },
      List: [],                 // JournalService.List — empty run history
      SignIns: [{ id: 's1', userDisplayName: 'Alice Smith', status: 'success' }],
      SignInsFiltered: [{ id: 's1', userDisplayName: 'Alice Smith', errorCode: 50126 }],
      // ActionsService (ADR-008): the built-in catalog, one preview, one apply.
      Catalog: connected ? [
        { id: 'user.signIn', page: 'users', danger: 'write', available: true, backend: 'graph', fields: [
          { name: 'user', kind: 'user', required: true },
          { name: 'state', kind: 'choice', required: true, options: ['blocked', 'allowed'], default: 'blocked' }] },
        { id: 'user.revokeSessions', page: 'users', danger: 'destructive', confirmField: 'user', available: true, backend: 'graph',
          fields: [{ name: 'user', kind: 'user', required: true }] },
        { id: 'group.membership', page: 'groups', danger: 'write', available: true, backend: 'graph', fields: [
          { name: 'user', kind: 'user', required: true }, { name: 'group', kind: 'group', required: true }] },
        { id: 'license.assign', page: 'licensing', danger: 'write', available: true, backend: 'graph', fields: [
          { name: 'user', kind: 'user', required: true }, { name: 'sku', kind: 'sku', required: true }] },
        { id: 'teams.effectivePolicies', page: 'teams', danger: 'read', available: true, backend: 'teams-ps',
          fields: [{ name: 'user', kind: 'user', required: true }] },
        { id: 'mailbox.fullAccess', page: 'users', danger: 'write', available: false,
          reason: { key: 'backendMissing', params: { backend: 'pwsh' } }, fields: null },
        // Go marshals an action without inputs as fields: null.
        { id: 'report.guests', page: 'reports', danger: 'read', available: true, backend: 'graph', fields: null, fanOut: true },
        { id: 'pack.contoso.hold', page: 'users', danger: 'write', available: false, pack: 'contoso',
          label: { en: 'Put a mailbox on hold' }, reason: { key: 'packUntrusted' }, fields: [] },
      ] : [],
      Plan: { id: 'p1', actionId: 'user.signIn', backend: 'graph', inputs: {}, changes: [
        { target: 'ann@contoso.com', field: 'signIn', op: 'set', before: 'allowed', after: 'blocked' }] },
      Run: { columns: ['policyType', 'policy'], backend: 'teams-ps', rows: [
        { policyType: 'meeting', policy: 'NoRecording' }, { policyType: 'calling', policy: 'Global' }] },
      Apply: { opId: 'o1', applied: 1, skipped: 0, failed: 0, canceled: false, outcomes: [
        { target: 'ann@contoso.com', field: 'signIn', op: 'set', ok: true, skipped: false }] },
    }
    // A test can override single bindings through window.__stub.
    const method = (name: string) => () => Promise.resolve((window as any).__stub?.[name] ?? results[name] ?? null)
    const service = new Proxy({}, { get: (_t, m: string) => method(m) })
    const namespace = new Proxy({}, { get: () => service })
    ;(window as any).go = new Proxy({}, { get: () => namespace })
    ;(window as any).runtime = new Proxy({}, {
      get: (_t, k: string) => (k === 'EventsOnMultiple' || k === 'EventsOn' ? () => () => {} : () => {}),
    })
  }, !!opts.connected)
}

test('renders without crashing and shows the connect page', async ({ page }) => {
  await stubWails(page)
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')

  await expect(page.getByText('SwissKnife', { exact: false }).first()).toBeVisible()
  await expect(page.getByText('Connect to a tenant')).toBeVisible()
  expect(errors).toEqual([])
})

test('sidebar navigation works and data tabs are locked while disconnected', async ({ page }) => {
  await stubWails(page)
  await page.goto('/')

  await page.getByRole('button', { name: 'Settings' }).click()
  await expect(page.getByText('Language')).toBeVisible()

  // Data pages must be disabled until a tenant connection exists.
  await expect(page.getByRole('button', { name: 'Raw Graph' })).toBeDisabled()
  await expect(page.getByRole('button', { name: 'Users & Admin' })).toBeDisabled()
})

test('grouped sidebar renders and run history opens without a connection', async ({ page }) => {
  await stubWails(page)
  await page.goto('/')

  // Section headers of the grouped navigation are visible.
  await expect(page.getByRole('button', { name: 'Insights' })).toBeVisible()
  // History is a local page: reachable while disconnected.
  await page.getByRole('button', { name: 'Run history' }).click()
  await expect(page.getByText('No runs recorded yet.')).toBeVisible()
})

test('task palette finds a task by words and opens its form on the page', async ({ page }) => {
  await stubWails(page, { connected: true })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')

  // Both entry points: the sidebar button and the Ctrl+K shortcut.
  const opener = page.getByRole('button', { name: /What do you need to do/ })
  await expect(opener).toBeVisible()
  await opener.click()
  const input = page.getByPlaceholder('Describe the task', { exact: false })
  await expect(input).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(input).toBeHidden()
  await page.keyboard.press('Control+k')
  await input.fill('private channel')
  await expect(page.getByRole('button', { name: /Add a user to a private channel/ })).toBeVisible()
  await page.keyboard.press('Enter')

  // Landed on Teams with that action's form already open.
  await expect(page.getByRole('heading', { name: 'Teams' })).toBeVisible()
  await expect(page.getByRole('heading', { name: 'Add someone to a private channel' })).toBeVisible()
  expect(errors).toEqual([])
})

// Tile labels are human text: escape them before they become a matcher, or the
// first label with a bracket in it silently matches the wrong element.
const label = (text: string) => new RegExp(text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))

test('Teams is action-first: every action is a visible tile with a hint', async ({ page }) => {
  await stubWails(page, { connected: true })
  await page.goto('/')
  await page.getByRole('button', { name: 'Teams', exact: true }).click()

  // All six capabilities are on screen without opening anything.
  for (const tile of [
    'Add someone to a team',
    'Add someone to a private channel',
    'Where is this person?',
    "What's inside a team",
    'Create a channel',
    'Turn a group into a team',
  ]) {
    await expect(page.getByRole('button', { name: label(tile) })).toBeVisible()
  }
  await expect(page.getByText('the first step before any private channel', { exact: false })).toBeVisible()

  // The view toggle is present; with nothing fetched the tiles keep the page.
  await page.getByRole('button', { name: 'Data', exact: true }).click()
  await expect(page.getByRole('button', { name: /Add someone to a team/ })).toBeVisible()
})

test('every migrated page renders its action tiles', async ({ page }) => {
  await stubWails(page, { connected: true })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')

  // One representative tile per migrated page: a crash or a missing i18n key in
  // any of them shows up here.
  const pages: [string, string][] = [
    ['Users & Admin', 'Everything about one user'],
    ['Licensing', 'Assign or remove a license'],
    ['Admin roles', 'Grant or revoke a role'],
    ['Groups', 'Add or remove group members'],
    ['App registrations', 'Secrets about to expire'],
    ['Chats', 'Who is in a chat'],
    ['Mail & Calendar', 'Send mail as a user'],
    ['Files', 'Browse a drive'],
    ['Intune', 'Wipe a device'],
    ['Entra devices', 'BitLocker recovery key'],
    ['Audit', 'Who changed what'],
    ['Service health', 'Active incidents'],
    ['Playbooks', 'Onboard a new employee'],
    ['Offboarding', "Copy someone's OneDrive to another account"],
    ['Cleanup', 'Find duplicate files'],
    ['Bulk / CSV', 'Create users from a list'],
    ['Usage reports', 'Who eats the OneDrive storage'],
    ['Security review', 'Review Conditional Access'],
  ]
  for (const [nav, tile] of pages) {
    await page.getByRole('button', { name: nav, exact: true }).click()
    await expect(page.getByRole('button', { name: label(tile) })).toBeVisible()
  }
  expect(errors).toEqual([])
})

test('"same access as" reaches the mirror form from the palette', async ({ page }) => {
  await stubWails(page, { connected: true })
  await page.goto('/')
  await page.getByRole('button', { name: /What do you need to do/ }).click()
  await page.getByPlaceholder('Describe the task', { exact: false }).fill('same access')
  await expect(page.getByRole('button', { name: /Give someone the same access as another user/ })).toBeVisible()
  await page.keyboard.press('Enter')

  // The copy form opens with both sides and the per-kind selection.
  await expect(page.getByRole('heading', { name: 'Give the same access as another user' })).toBeVisible()
  await expect(page.getByText('Copy access from (source)')).toBeVisible()
  await expect(page.getByText('Give it to (target)')).toBeVisible()
  await expect(page.getByRole('button', { name: /Preview the difference/ })).toBeVisible()
})

test('a result can be cleared without leaving the page', async ({ page }) => {
  await stubWails(page, { connected: true })
  await page.goto('/')
  await page.getByRole('button', { name: 'Audit', exact: true }).click()

  // Running an action shows the data pane and the way back out of it.
  await page.getByRole('button', { name: /Sign-in logs \(whole tenant\)/ }).click()
  await page.getByRole('button', { name: 'Run', exact: true }).click()
  const clear = page.getByRole('button', { name: 'Clear result' })
  await expect(clear).toBeVisible()

  await clear.click()
  await expect(clear).toBeHidden()
  await expect(page.getByRole('button', { name: /Who changed what/ })).toBeVisible()
})

test('language switch to Russian localizes the UI', async ({ page }) => {
  await stubWails(page)
  await page.goto('/')

  await page.getByRole('button', { name: 'Settings' }).click()
  // The language selector is the first <select> on the settings page.
  await page.locator('select').first().selectOption('ru')
  await expect(page.getByText('Настройки').first()).toBeVisible()
})

test('a catalog action previews its change before applying it', async ({ page }) => {
  await stubWails(page, { connected: true })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  await page.getByRole('button', { name: 'Users & Admin' }).click()

  // An action whose backend is missing stays visible and says why.
  await expect(page.getByText('Needs the pwsh backend, which is not set up.')).toBeVisible()
  // A community pack's action carries its own label and says why it is off.
  await expect(page.getByText('Put a mailbox on hold')).toBeVisible()
  await expect(page.getByText('not trusted yet', { exact: false })).toBeVisible()

  await page.getByRole('button', { name: /Block or unblock sign-in/ }).click()
  await page.getByRole('button', { name: 'User', exact: true }).click()
  await page.getByPlaceholder('Search or paste', { exact: false }).fill('ann@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: 'Preview changes' }).click()

  await expect(page.getByText('What will change')).toBeVisible()
  await expect(page.getByText('allowed →', { exact: false })).toBeVisible()
  await page.getByRole('button', { name: 'Apply', exact: true }).click()
  await expect(page.getByText('Done: 1 changed, 0 already in place').first()).toBeVisible()
  expect(errors).toEqual([])
})

test('a read action shows its rows in the tile', async ({ page }) => {
  await stubWails(page, { connected: true })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  await page.getByRole('button', { name: 'Teams', exact: true }).click()

  const tile = page.getByRole('button', { name: /Which Teams policies apply to a person/ })
  await expect(tile).toContainText('Teams PS') // backend badge
  await tile.click()
  await page.getByRole('button', { name: 'User', exact: true }).click()
  await page.getByPlaceholder('Search or paste', { exact: false }).fill('bob@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: 'Show', exact: true }).click()
  await expect(page.getByRole('cell', { name: 'NoRecording' })).toBeVisible()
  await expect(page.getByRole('cell', { name: 'Meetings' })).toBeVisible() // translated value
  expect(errors).toEqual([])
})

test('the explorer opens an action with the picked object filled in', async ({ page }) => {
  await stubWails(page, { connected: true })
  await page.addInitScript(() => {
    ;(window as any).__stub = { List: [{ id: 'u1', displayName: 'Ann Lee', userPrincipalName: 'ann@contoso.com' }] }
  })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  await page.getByRole('button', { name: 'Explorer', exact: true }).click()
  await page.getByRole('button', { name: /Ann Lee/ }).click()
  await page.getByRole('button', { name: /Block or unblock sign-in/ }).click()

  // The Users page opens with the action's panel and the user already chosen.
  await expect(page.getByRole('button', { name: 'Preview changes' })).toBeEnabled()
  await page.getByRole('button', { name: 'Preview changes' }).click()
  await expect(page.getByText('What will change')).toBeVisible()
  expect(errors).toEqual([])
})

test('on-prem AD works without a tenant connection', async ({ page }) => {
  await stubWails(page)
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  // Hidden until switched on: the connect page offers it.
  await expect(page.getByRole('button', { name: 'On-prem AD', exact: true })).toHaveCount(0)
  await page.getByRole('button', { name: 'Turn on On-prem AD' }).click()
  await expect(page.getByRole('button', { name: 'On-prem AD', exact: true })).toBeVisible()
  await page.getByRole('button', { name: /Directory connection/ }).click()
  await expect(page.getByText('Not connected to a directory.')).toBeVisible()
  await page.getByRole('combobox').selectOption('none')
  await expect(page.getByText('crosses the network in clear text', { exact: false })).toBeVisible()
  expect(errors).toEqual([])
})

test('an on-prem-only setup hides the Microsoft 365 pages and starts on AD', async ({ page }) => {
  await stubWails(page)
  await page.addInitScript(() => localStorage.setItem('workspaces', JSON.stringify({ cloud: false, onprem: true })))
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  await expect(page.getByRole('button', { name: /Directory connection/ })).toBeVisible()
  await expect(page.getByRole('button', { name: 'Connect', exact: true })).toHaveCount(0)
  await expect(page.getByRole('button', { name: 'Users & Admin' })).toHaveCount(0)
  await expect(page.getByRole('button', { name: 'Settings' })).toBeVisible()
  expect(errors).toEqual([])
})

// Every screen, every tile, connected and with both workspaces on: a page
// that crashes (blank window) or throws shows up here.
test('every screen opens and every tile can be clicked without crashing', async ({ page }) => {
  test.setTimeout(180_000)
  await stubWails(page, { connected: true })
  await page.addInitScript(() => localStorage.setItem('workspaces', JSON.stringify({ cloud: true, onprem: true })))
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  const names = await page.locator('nav button').evaluateAll((els) =>
    els.filter((e) => !e.hasAttribute('aria-expanded')).map((e) => (e.textContent || '').trim()))
  expect(names.length).toBeGreaterThan(20)
  expect(names).toContain('Usage reports')
  const crashed = page.getByText('This page ran into a problem')
  for (const name of names) {
    // Catalog tiles arrive after the page: wait until they are in.
    const open = async () => {
      await page.locator('nav button').filter({ hasText: name }).first().click()
      await expect(page.locator('main h1').first(), name).toBeVisible()
      await expect(page.locator('main [data-catalog-ready="false"]'), name).toHaveCount(0)
    }
    await open()
    await expect(crashed, name).toHaveCount(0)
    const tiles = await page.locator('main [data-tile]').evaluateAll((els) =>
      els.filter((e) => e.getAttribute('aria-disabled') !== 'true').map((e) => e.getAttribute('data-tile')))
    for (const id of tiles) {
      await page.locator(`main [data-tile="${id}"]`).first().click()
      await expect(page.locator('main h1').first(), `${name} → ${id}`).toBeVisible()
      await expect(crashed, `${name} → ${id}`).toHaveCount(0)
      expect(errors, `${name} → ${id}`).toEqual([])
      await open() // back to a clean page (a tile may navigate or open a panel)
    }
  }
  expect(errors).toEqual([])
})

test('the menu folds to icons, hides pages, and support is one click away', async ({ page }) => {
  await stubWails(page, { connected: true })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')

  // Icons only: names come back as a hover label.
  await page.getByRole('button', { name: 'Collapse the menu' }).click()
  await expect(page.locator('nav').getByText('Licensing')).toHaveCount(0)
  await page.locator('nav').getByRole('button', { name: 'Licensing' }).hover()
  await expect(page.getByRole('tooltip')).toHaveText('Licensing')
  await page.getByRole('button', { name: 'Expand the menu' }).click()
  await expect(page.locator('nav').getByText('Licensing')).toBeVisible()

  // Essentials only: rarely used pages leave the menu, Ctrl+K still finds them.
  await page.locator('nav').getByRole('button', { name: 'Settings' }).click()
  await page.getByRole('button', { name: 'Essentials only' }).click()
  await expect(page.locator('nav').getByRole('button', { name: 'Raw Graph', exact: false })).toHaveCount(0)
  await expect(page.locator('nav').getByRole('button', { name: 'Settings' })).toBeVisible()
  await page.getByRole('button', { name: 'Show everything' }).click()

  await expect(page.getByRole('button', { name: /Support on Tribute/ })).toBeVisible()
  expect(errors).toEqual([])
})

test('capabilities show what runs, the fallbacks and why the others cannot', async ({ page }) => {
  await stubWails(page, { connected: true })
  await page.addInitScript(() => {
    ;(window as any).__stub = { Capabilities: [
      { capability: 'exchange.mailbox.sendOnBehalf', action: 'mailbox.sendOnBehalf', page: 'mail', danger: 'write', impls: [
        { backend: 'exo-api', state: 'unavailable', reason: { key: 'exoApiNotEnabled' } },
        { backend: 'exo-ps', state: 'runs', via: 'worker' }] },
      { capability: 'teams.policy.list', action: 'teams.policies', page: 'teams', danger: 'read', impls: [
        { backend: 'teams-ps', state: 'unavailable', reason: { key: 'pwshMissing' } }] },
      { capability: 'entra.user.signIn', action: 'user.signIn', page: 'users', danger: 'write', impls: null },
    ] }
  })
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(String(e)))
  await page.goto('/')
  await page.locator('nav').getByRole('button', { name: 'Settings' }).click()
  await page.getByRole('button', { name: 'Show capabilities' }).click()

  const mail = page.locator('[data-capability="exchange.mailbox.sendOnBehalf"]')
  await expect(mail.getByText('Exchange PS · via worker')).toBeVisible()
  await expect(mail.getByText(/Exchange: The Exchange Admin API is not available/)).toBeVisible()
  await expect(page.getByText('1 of 3 can run now')).toBeVisible()
  await page.getByLabel('Only those that cannot run').check()
  await expect(mail).toHaveCount(0)
  await expect(page.locator('[data-capability="teams.policy.list"]').getByText(/Needs PowerShell 7/)).toBeVisible()
  expect(errors).toEqual([])
})
