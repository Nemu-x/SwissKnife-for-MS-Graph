import { defineConfig } from 'vitest/config'

// Unit tests live under src/ as *.test.ts. The e2e/ folder is Playwright's
// (*.spec.ts) and must not be picked up by vitest's default glob.
export default defineConfig({
  test: {
    include: ['src/**/*.test.{ts,tsx}'],
  },
})
