import { createContext, useContext, useEffect, useState, type ReactNode, useCallback } from 'react'
import { api, type Status } from './api'
import { applyAccent } from './color'

type Theme = 'dark' | 'light'
type Toast = { id: number; kind: 'ok' | 'err' | 'info'; text: string }

interface Store {
  status: Status | null
  setStatus: (s: Status | null) => void
  refreshStatus: () => Promise<void>
  connected: boolean
  readOnly: boolean

  domains: string[]
  loadDomains: () => Promise<void>

  theme: Theme
  toggleTheme: () => void

  accent: string
  setAccent: (hex: string) => void

  safeMode: boolean
  setSafeMode: (v: boolean) => void

  toasts: Toast[]
  toast: (kind: Toast['kind'], text: string) => void
  dismiss: (id: number) => void
}

const Ctx = createContext<Store | null>(null)

export function StoreProvider({ children }: { children: ReactNode }) {
  const [status, setStatus] = useState<Status | null>(null)
  const [theme, setTheme] = useState<Theme>((localStorage.getItem('theme') as Theme) || 'dark')
  const [safeMode, setSafeModeState] = useState<boolean>(localStorage.getItem('safeMode') !== 'false')
  const [toasts, setToasts] = useState<Toast[]>([])
  const [domains, setDomains] = useState<string[]>([])
  // '' = Auto (use the theme palette accent); a hex overrides it.
  const [accent, setAccentState] = useState<string>(localStorage.getItem('accent') ?? '')

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme)
    localStorage.setItem('theme', theme)
  }, [theme])

  // Re-apply accent whenever it changes or the theme flips (theme resets vars).
  useEffect(() => {
    applyAccent(accent)
    localStorage.setItem('accent', accent)
  }, [accent, theme])

  const toast = useCallback((kind: Toast['kind'], text: string) => {
    const id = Date.now() + Math.random()
    setToasts((t) => [...t, { id, kind, text }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 5000)
  }, [])

  const refreshStatus = useCallback(async () => {
    try {
      setStatus(await api.connect.status())
    } catch {
      /* ignore */
    }
  }, [])

  useEffect(() => {
    refreshStatus()
  }, [refreshStatus])

  const loadDomains = useCallback(async () => {
    try {
      setDomains(await api.connect.domains())
    } catch {
      setDomains([])
    }
  }, [])

  // Clear cached domains on disconnect; loading is opt-in (checkbox on Connect).
  useEffect(() => {
    if (!status?.connected) setDomains([])
  }, [status?.connected])

  const value: Store = {
    status,
    setStatus,
    refreshStatus,
    connected: !!status?.connected,
    readOnly: !!status?.readOnly,
    domains,
    loadDomains,
    theme,
    toggleTheme: () => setTheme((t) => (t === 'dark' ? 'light' : 'dark')),
    accent,
    setAccent: setAccentState,
    safeMode,
    setSafeMode: (v) => {
      setSafeModeState(v)
      localStorage.setItem('safeMode', String(v))
    },
    toasts,
    toast,
    dismiss: (id) => setToasts((t) => t.filter((x) => x.id !== id)),
  }

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>
}

export function useStore(): Store {
  const s = useContext(Ctx)
  if (!s) throw new Error('useStore must be used within StoreProvider')
  return s
}
