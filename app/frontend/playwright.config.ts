import { defineConfig } from '@playwright/test'

// GUI smoke: serves the production build (dist/) with stubbed Wails bindings.
// Run `npm run build` first; `npm run e2e` handles both.
export default defineConfig({
  testDir: './e2e',
  timeout: 30_000,
  use: {
    baseURL: 'http://localhost:4173',
  },
  webServer: {
    command: 'npx vite preview --port 4173 --strictPort',
    url: 'http://localhost:4173',
    reuseExistingServer: true,
    timeout: 30_000,
  },
})
