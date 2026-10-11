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

const user = { name: 'user', kind: 'user', required: true }
const yesNo = (name: string, label: string) => ({ name, kind: 'choice', options: ['yes', 'no'], default: 'yes', label: { en: label } })

const CATALOG = [
  { id: 'pack.security-basics.compromisedUser', page: 'security', danger: 'destructive', confirmField: 'user', available: true, backend: 'workflow',
    pack: 'security-basics', workflow: true, capability: 'pack.security-basics.compromisedUser',
    label: { en: 'Compromised account response' }, hint: { en: 'Block, sign out everywhere, then clean up what an attacker leaves behind.' },
    fields: [{ ...user, label: { en: 'Account' } }, yesNo('resetPassword', 'Reset the password'), yesNo('resetMfa', 'Reset MFA'), yesNo('cleanMailbox', 'Stop forwarding and suspicious rules')] },
  { id: 'teams.userPolicy', page: 'teams', danger: 'write', available: true, backend: 'teams-ps', fields: [user,
    { name: 'policyType', kind: 'choice', required: true, options: ['meeting', 'messaging', 'calling', 'appSetup', 'appPermission'], default: 'meeting' },
    { name: 'policy', kind: 'text', required: true, default: 'Global' }] },
  { id: 'teams.groupPolicy', page: 'teams', danger: 'write', available: true, backend: 'teams-ps', fields: [{ name: 'group', kind: 'group', required: true },
    { name: 'policyType', kind: 'choice', required: true, options: ['meeting', 'messaging', 'calling', 'appSetup', 'appPermission'], default: 'meeting' },
    { name: 'policy', kind: 'text', required: true, default: 'Global' }] },
  { id: 'teams.policies', page: 'teams', danger: 'read', available: true, backend: 'teams-ps', fields: [] },
  { id: 'teams.effectivePolicies', page: 'teams', danger: 'read', available: true, backend: 'teams-ps', fields: [user] },
]

// What the compromised-account workflow would change, grouped by step.
const PLAN = {
  id: 'p1', actionId: 'pack.security-basics.compromisedUser', backend: 'workflow', inputs: {}, confirmTarget: 'nadia@contoso.com', changes: [
    { target: 'nadia@contoso.com', field: 'signIn', op: 'set', before: 'allowed', after: 'blocked', step: 'user.signIn', stepNo: 1 },
    { target: 'nadia@contoso.com', field: 'sessions', op: 'set', before: 'active', after: 'revoked', step: 'user.revokeSessions', stepNo: 2 },
    { target: 'nadia@contoso.com', field: 'password', op: 'set', after: 'randomPassword', step: 'user.resetPassword', stepNo: 3 },
    { target: 'nadia@contoso.com', field: 'authMethod', op: 'remove', before: 'Microsoft Authenticator (iPhone 15)', step: 'user.resetMfa', stepNo: 4 },
    { target: 'nadia@contoso.com', field: 'authMethod', op: 'remove', before: 'phone +1 555 0100', step: 'user.resetMfa', stepNo: 4 },
    { target: 'nadia@contoso.com', field: 'inboxRule', op: 'set', before: 'enabled', after: 'disabled', note: '', step: 'mail.disableRules', stepNo: 5 },
    { target: 'nadia@contoso.com', field: 'forwarding', op: 'remove', before: 'smtp:billing@contoso-invoices.example', step: 'mailbox.forwarding', stepNo: 6 },
  ],
}

const pack = (name: string, category: string, kind: string, title: string, description: string, installed = '') =>
  ({ name, version: '1.0.0', category, kind, title: { en: title }, description: { en: description }, author: 'SwissKnife project',
    signed: true, installed, update: false, digest: name })
const HUB = {
  url: 'https://raw.githubusercontent.com/Nemu-x/SwissKnife-Action-Hub/main/', custom: false, packs: [
    pack('security-basics', 'security', 'workflow', 'Security basics', 'Compromised account response, shared-mailbox conversion and a quick offboarding — built-in actions only.', '1.1.0'),
    pack('phishing-response', 'security', 'workflow', 'Phishing response', 'Block the sender, report the message to Microsoft and purge it from every mailbox — or undo a false positive.'),
    pack('hr-lifecycle', 'hr', 'workflow', 'HR lifecycle', 'New joiner, team move, leave of absence and return from leave — built-in actions only.'),
    pack('mailbox-delegation', 'exchange', 'workflow', 'Mailbox delegation', "Give or take back an assistant's access in one step, open a shared mailbox to a user, set up a room."),
    pack('exchange-audit', 'reports', 'script', 'Exchange audit reports', 'Who can open each shared mailbox, mailboxes with auditing off, distribution lists nobody is in (PowerShell, read-only).'),
    pack('teams-governance', 'teams', 'workflow', 'Teams governance', 'Give a department or a person its meeting, messaging and calling policies in one step.'),
    pack('intune-rollout', 'intune', 'workflow', 'Intune rollout', 'Roll a configuration profile and a compliance policy out to a group together — or take both back.'),
  ],
}

const impl = (backend: string, state: string, extra: object = {}) => ({ backend, state, ...extra })
const CAPS = [
  { capability: 'entra.user.revokeSessions', action: 'user.revokeSessions', page: 'users', danger: 'destructive', impls: [impl('graph', 'runs')] },
  { capability: 'entra.user.signIn', action: 'user.signIn', page: 'users', danger: 'write', impls: [impl('graph', 'runs')] },
  { capability: 'exchange.mailbox.folderPermission', action: 'mailbox.folderPermission', page: 'mail', danger: 'write',
    impls: [impl('exo-api', 'runs'), impl('exo-ps', 'ready', { via: 'worker' })] },
  { capability: 'exchange.mailbox.fullAccess', action: 'mailbox.fullAccess', page: 'mail', danger: 'write', impls: [impl('exo-ps', 'runs', { via: 'worker' })] },
  { capability: 'exchange.mailbox.sendOnBehalf', action: 'mailbox.sendOnBehalf', page: 'mail', danger: 'write',
    impls: [impl('exo-api', 'runs'), impl('exo-ps', 'ready', { via: 'worker' })] },
  { capability: 'exchange.transportRule.state', action: 'transportRule.state', page: 'mail', danger: 'write', impls: [impl('exo-ps', 'runs', { via: 'worker' })] },
  { capability: 'ad.user.unlock', action: 'ad.unlock', page: 'onprem', danger: 'write', impls: [impl('ldap', 'unavailable', { reason: { key: 'ldapNotConnected' } })] },
  { capability: 'teams.policy.assignUser', action: 'teams.userPolicy', page: 'teams', danger: 'write', impls: [impl('teams-ps', 'runs', { via: 'local' })] },
]

async function stub(page: Page) {
  await page.addInitScript(([users, catalog, plan, hub, caps]: unknown[]) => {
    // Bindings are window.go.<package>.<Service>.<Method>: answers can be
    // keyed "Service.Method" when two services share a method name.
    const results: Record<string, unknown> = {
      GetStatus: { connected: true, profileName: 'contoso', readOnly: false },
      Profiles: [], Domains: ['contoso.com'], Version: '2.0.0', Probe: {},
      Check: { currentVersion: '2.0.0', latestVersion: '2.0.0', updateAvailable: false, notes: '', url: '' },
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
      'UsersService.List': users,
      'PacksService.List': [],
      'JournalService.List': [],
      Catalog: catalog,
      Plan: plan,
      HubCatalog: hub,
      Capabilities: caps,
      Keys: [],
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
    }
    const service = (svc: string) => new Proxy({}, {
      get: (_t, m: string) => () => Promise.resolve(results[`${svc}.${m}`] ?? results[m] ?? null),
    })
    const namespace = new Proxy({}, { get: (_t, svc: string) => service(svc) })
    ;(window as any).go = new Proxy({}, { get: () => namespace })
    ;(window as any).runtime = new Proxy({}, {
      get: (_t, k: string) => (k === 'EventsOnMultiple' || k === 'EventsOn' ? () => () => {} : () => {}),
    })
  }, [USERS, CATALOG, PLAN, HUB, CAPS])
}

// shoot brings a card to the top of the viewport and captures the window.
async function shoot(page: Page, file: string) {
  await page.waitForTimeout(400)
  await page.screenshot({ path: `${OUT}/${file}` })
}

test('README screenshots', async ({ page }) => {
  await page.setViewportSize({ width: 1360, height: 900 })
  await stub(page)
  await page.goto('/')
  await shoot(page, 'dashboard.png')

  // The palette: how the app is meant to be entered.
  await page.getByRole('button', { name: /What do you need to do/ }).click()
  await page.getByPlaceholder('Describe the task', { exact: false }).fill('access')
  await shoot(page, 'palette.png')
  await page.keyboard.press('Escape')

  // A workflow pack: one preview, grouped by step, one confirmation.
  await page.locator('nav').getByRole('button', { name: /Security/ }).first().click()
  await page.getByRole('button', { name: /Compromised account response/ }).click()
  await page.getByRole('button', { name: 'Account', exact: true }).click()
  await page.getByPlaceholder('Search or paste', { exact: false }).fill('nadia@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: 'Preview changes' }).click()
  // Tall enough for every step of the plan.
  await page.setViewportSize({ width: 1360, height: 1180 })
  await page.getByText('What will change').evaluate((el) => el.closest('section, form, div')?.scrollIntoView({ block: 'center' }))
  await shoot(page, 'workflow-preview.png')
  await page.setViewportSize({ width: 1360, height: 900 })

  // Exchange and Teams PowerShell as tiles.
  await page.locator('nav').getByRole('button', { name: 'Teams', exact: true }).click()
  await shoot(page, 'teams.png')

  // Settings: the Action Hub catalog and the capabilities view.
  await page.locator('nav').getByRole('button', { name: 'Settings' }).click()
  const packs = page.locator('section').filter({ has: page.getByRole('heading', { name: 'Action packs' }) })
  await packs.getByRole('button', { name: 'Action Hub' }).click()
  await packs.getByText('Phishing response').waitFor()
  await packs.evaluate((el) => el.scrollIntoView({ block: 'start' }))
  await shoot(page, 'action-hub.png')

  const caps = page.locator('section').filter({ has: page.getByRole('heading', { name: 'Capabilities' }) })
  await caps.getByRole('button', { name: 'Show capabilities' }).click()
  await caps.locator('[data-capability]').first().waitFor()
  await caps.evaluate((el) => el.scrollIntoView({ block: 'start' }))
  await shoot(page, 'capabilities.png')

  // Access mirror with a real-looking diff.
  await page.locator('nav').getByRole('button', { name: 'Users & Admin' }).click()
  await page.getByRole('button', { name: /Give the same access as another user/ }).click()
  await page.getByRole('button', { name: 'Copy access from (source)' }).first().click()
  await page.getByPlaceholder('Search or paste an ID…').fill('alice@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: 'Give it to (target)' }).first().click()
  await page.getByPlaceholder('Search or paste an ID…').fill('marcus@contoso.com')
  await page.keyboard.press('Enter')
  await page.getByRole('button', { name: /Preview the difference/ }).click()
  await shoot(page, 'access-mirror.png')
})
