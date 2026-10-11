import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// Content-Security-Policy of the shipped bundle. 'self' covers the Wails
// asset origin (http://wails.localhost on Windows, wails://wails elsewhere)
// and the runtime scripts Wails injects from it (/wails/ipc.js, runtime.js).
// Inline styles stay allowed (React style props, Tailwind); scripts may only
// come from the bundle. ws:/localhost cover `wails dev` serving a build.
const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: https:",
  "font-src 'self' data:",
  "connect-src 'self' ws: http://localhost:*",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'none'",
].join('; ')

// The meta tag is added to production builds only: the Vite dev server injects
// an inline React-refresh preamble that script-src 'self' would block.
function contentSecurityPolicy(): Plugin {
  return {
    name: 'content-security-policy',
    apply: 'build',
    transformIndexHtml: () => [
      { tag: 'meta', attrs: { 'http-equiv': 'Content-Security-Policy', content: CSP }, injectTo: 'head-prepend' },
    ],
  }
}

export default defineConfig({
  plugins: [react(), tailwindcss(), contentSecurityPolicy()],
})
