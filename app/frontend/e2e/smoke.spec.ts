import { test, expect, type Page } from '@playwright/test'

// The app normally runs inside Wails, which injects window.go (bindings) and
// window.runtime (events). We stub both so the real production bundle can be
// exercised in a plain browser: any binding call resolves with a sane default.
// This catches white-screen crashes, i18n init breakage, and nav regressions.
async function stubWails(page: Page) {
  await page.addInitScript(() => {
    const results: Record<string, unknown> = {
      GetStatus: null,          // not connected
      Profiles: [],
      Domains: [],
      Version: '0.0.0-e2e',
      Probe: {},
      Check: { currentVersion: '0.0.0-e2e', latestVersion: '', updateAvailable: false, notes: '', url: '' },
    }
    const method = (name: string) => () => Promise.resolve(results[name] ?? null)
    const service = new Proxy({}, { get: (_t, m: string) => method(m) })
    const namespace = new Proxy({}, { get: () => service })
    ;(window as any).go = new Proxy({}, { get: () => namespace })
    ;(window as any).runtime = new Proxy({}, {
      get: (_t, k: string) => (k === 'EventsOnMultiple' || k === 'EventsOn' ? () => () => {} : () => {}),
    })
  })
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

test('language switch to Russian localizes the UI', async ({ page }) => {
  await stubWails(page)
  await page.goto('/')

  await page.getByRole('button', { name: 'Settings' }).click()
  // The language selector is the first <select> on the settings page.
  await page.locator('select').first().selectOption('ru')
  await expect(page.getByText('Настройки').first()).toBeVisible()
})
