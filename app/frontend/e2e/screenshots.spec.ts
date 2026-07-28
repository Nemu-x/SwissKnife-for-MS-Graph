import { test, type Page } from '@playwright/test'
import path from 'node:path'

// README screenshots, generated from the production bundle with stubbed Wails
// bindings — no tenant and no real data involved. Opt in explicitly:
//   SHOTS=1 npx playwright test e2e/screenshots.spec.ts
// The files land in docs/screenshots/ and are committed with the README.
// Playwright runs from app/frontend; the repo root is two levels up.
const OUT = path.resolve(process.cwd(), '../../docs/screenshots')

test.skip(!process.env.SHOTS, 'set SHOTS=1 to regenerate the README screenshots')

const USERS = [
  { id: 'u1', displayName: 'Alice Whitfield', userPrincipalName: 'alice@contoso.com', accountEnabled: true, jobTitle: 'Support Lead' },
  { id: 'u2', displayName: 'Marcus Bell', userPrincipalName: 'marcus@contoso.com', accountEnabled: true, jobTitle: 'Analyst' },
  { id: 'u3', displayName: 'Nadia Fisher', userPrincipalName: 'nadia@contoso.com', accountEnabled: false, jobTitle: 'Contractor' },
]

async function stub(page: Page) {
  await page.addInitScript(([users]: [unknown]) => {
    const results: Record<string, unknown> = {
      GetStatus: { connected: true, profileName: 'contoso', readOnly: false },
      Profiles: [], Domains: ['contoso.com'], Version: '1.0.0', Probe: {},
      Check: { currentVersion: '1.0.0', latestVersion: '1.0.0', updateAvailable: false, notes: '', url: '' },
      Summary: {
        orgName: 'Contoso Ltd', users: 214, groups: 38, domains: 3, licensesUsed: 187,
        licenses: [
          { skuPartNumber: 'SPE_E3', consumed: 120, total: 150 },
          { skuPartNumber: 'ENTERPRISEPACK', consumed: 52, total: 55 },
          { skuPartNumber: 'EXCHANGESTANDARD', consumed: 15, total: 40 },
          { skuPartNumber: 'FLOW_FREE', consumed: 96, total: 10000 },
          { skuPartNumber: 'TEAMS_EXPLORATORY', consumed: 12, total: 100 },
        ],
      },
      List: users,
      Compare: [
        { kind: 'group', id: 'g1', name: 'Global Finance', status: 'missing', copyable: true },
        { kind: 'group', id: 'g2', name: 'IT_Helpdesk', status: 'missing', copyable: true },
        { kind: 'team', id: 't2', name: 'Intermark Team', status: 'missing', copyable: true },
        { kind: 'channel', id: 'c1', name: 'Escalations', teamId: 't1', teamName: 'HelpCenter', sub: 'private', status: 'missing', copyable: true },
        { kind: 'group', id: 'dl', name: 'All Company', status: 'missing', copyable: false, reasonKey: 'reasons.exchangeGroup' },
        { kind: 'license', id: 'sku', name: 'SPE_E3', status: 'both', copyable: false },
        { kind: 'group', id: 'g9', name: 'Interns 2026', status: 'targetOnly', copyable: false },
      ],
      ListAllTeams: [
        { id: 't1', displayName: 'HelpCenter', description: 'First line support' },
        { id: 't2', displayName: 'Intermark Team' },
      ],
      CAPolicies: [
        {
          id: 'p1', displayName: 'Require MFA for all users', state: 'enabled',
          conditions: {
            users: { includeUsers: ['All'], excludeGroups: ['g-breakglass'] },
            applications: { includeApplications: ['All'] },
            clientAppTypes: ['browser', 'mobileAppsAndDesktopClients'],
          },
          grantControls: { operator: 'OR', builtInControls: ['mfa'] },
        },
        { id: 'p2', displayName: 'Block legacy authentication', state: 'enabled', conditions: { users: { includeUsers: ['All'] }, applications: { includeApplications: ['All'] }, clientAppTypes: ['exchangeActiveSync', 'other'] }, grantControls: { operator: 'OR', builtInControls: ['block'] } },
        { id: 'p3', displayName: 'Compliant device for admins', state: 'enabledForReportingButNotEnforced', conditions: { users: { includeRoles: ['r1', 'r2', 'r3'] }, applications: { includeApplications: ['All'] } }, grantControls: { operator: 'AND', builtInControls: ['mfa', 'compliantDevice'] } },
      ],
      Skus: [
        { skuId: 'sku1', skuPartNumber: 'SPE_E3', consumedUnits: 120, prepaidUnits: { enabled: 150 } },
        { skuId: 'sku2', skuPartNumber: 'ENTERPRISEPACK', consumedUnits: 52, prepaidUnits: { enabled: 55 } },
      ],
      SignInsFiltered: [
        { id: 's1', userDisplayName: 'Nadia Fisher', userPrincipalName: 'nadia@contoso.com', appDisplayName: 'Office 365 Exchange Online', ipAddress: '203.0.113.24', createdDateTime: '2026-07-28T09:14:02Z', errorCode: 53003, failureReason: 'Blocked by Conditional Access' },
        { id: 's2', userDisplayName: 'Nadia Fisher', userPrincipalName: 'nadia@contoso.com', appDisplayName: 'Microsoft Teams', ipAddress: '203.0.113.24', createdDateTime: '2026-07-28T09:11:47Z', errorCode: 50126, failureReason: 'Invalid username or password' },
      ],
    }
    const method = (name: string) => () => Promise.resolve(results[name] ?? null)
    const service = new Proxy({}, { get: (_t, m: string) => method(m) })
    const namespace = new Proxy({}, { get: () => service })
    ;(window as any).go = new Proxy({}, { get: () => namespace })
    ;(window as any).runtime = new Proxy({}, {
      get: (_t, k: string) => (k === 'EventsOnMultiple' || k === 'EventsOn' ? () => () => {} : () => {}),
    })
  }, [USERS] as [unknown])
}

test('README screenshots', async ({ page }) => {
  await page.setViewportSize({ width: 1360, height: 900 })
  await stub(page)
  await page.goto('/')
  await page.waitForTimeout(400)
  await page.screenshot({ path: `${OUT}/dashboard.png` })

  // The palette: how the app is meant to be entered.
  await page.getByRole('button', { name: /What do you need to do/ }).click()
  await page.getByPlaceholder('Describe the task', { exact: false }).fill('access')
  await page.waitForTimeout(300)
  await page.screenshot({ path: `${OUT}/palette.png` })
  await page.keyboard.press('Escape')

  // Action-first page: every capability visible, no drawers.
  await page.getByRole('button', { name: 'Teams', exact: true }).click()
  await page.waitForTimeout(300)
  await page.screenshot({ path: `${OUT}/teams.png` })

  await page.getByRole('button', { name: 'Users & Admin' }).click()
  await page.waitForTimeout(300)
  await page.screenshot({ path: `${OUT}/users.png` })

  // Access mirror with a real-looking diff.
  await page.getByRole('button', { name: /Give the same access as another user/ }).click()
  await page.getByRole('button', { name: 'Copy access from (source)' }).first().click()
  await page.getByPlaceholder('Search or paste an ID…').fill('alice@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: 'Give it to (target)' }).first().click()
  await page.getByPlaceholder('Search or paste an ID…').fill('marcus@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: /Preview the difference/ }).click()
  await page.waitForTimeout(500)
  await page.screenshot({ path: `${OUT}/access-mirror.png` })

  // Audit: the "why can this person not sign in" flow.
  await page.getByRole('button', { name: 'Audit', exact: true }).click()
  await page.getByRole('button', { name: /Why can this person not sign in/ }).click()
  await page.getByRole('button', { name: 'User (UPN or ID)' }).first().click()
  await page.getByPlaceholder('Search or paste an ID…').fill('nadia@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: /Failed sign-ins only/ }).first().click()
  await page.waitForTimeout(500)
  await page.screenshot({ path: `${OUT}/audit.png` })

  // Playbooks: the onboarding run with its role profile.
  await page.getByRole('button', { name: 'Playbooks' }).click()
  await page.getByRole('button', { name: /Onboard a new employee/ }).click()
  await page.waitForTimeout(400)
  await page.screenshot({ path: `${OUT}/playbooks.png` })

  // Conditional Access, read as facts instead of raw JSON.
  await page.getByRole('button', { name: 'Security review' }).click()
  await page.getByRole('button', { name: /Review Conditional Access/ }).click()
  await page.getByRole('button', { name: /Load policies/ }).click()
  await page.waitForTimeout(400)
  await page.getByRole('button', { name: /Require MFA for all users/ }).click()
  await page.waitForTimeout(300)
  await page.screenshot({ path: `${OUT}/security.png` })

  // Raw Graph console stays part of the story for the API-minded.
  await page.getByRole('button', { name: 'Raw Graph' }).click()
  await page.waitForTimeout(300)
  await page.screenshot({ path: `${OUT}/raw.png` })
})
