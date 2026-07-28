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
    }
    const method = (name: string) => () => Promise.resolve(results[name] ?? null)
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
    ['Groups', 'Add someone to a group'],
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
