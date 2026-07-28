import { useCallback, useEffect, useRef, useState } from 'react'
import { Layout, LOCAL_PAGES } from './components/Layout'
import { Toasts } from './components/Toasts'
import { CommandPalette } from './components/CommandPalette'
import { StoreProvider, useStore } from './lib/store'
import { pages, type PageId } from './pages/registry'

function Shell() {
  const [page, setPage] = useState<PageId>('connect')
  const [paletteOpen, setPaletteOpen] = useState(false)
  const { connected, requestAction, setNavigator } = useStore()
  const wasConnected = useRef(false)

  // Land on the dashboard right after a fresh connection.
  useEffect(() => {
    if (connected && !wasConnected.current) setPage('dashboard')
    wasConnected.current = connected
  }, [connected])

  // Navigating drops any unclaimed palette request, so a task that could not be
  // honoured on its own page never fires on the next one.
  const navigate = useCallback((p: PageId) => {
    requestAction(null)
    setPage(p)
  }, [requestAction])

  // Pages navigate through the store (dashboard tiles, cross-page links).
  useEffect(() => {
    setNavigator((page, action) => {
      navigate(page as PageId)
      if (action) requestAction(action)
    })
  }, [setNavigator, navigate, requestAction])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        setPaletteOpen((o) => !o)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  // if disconnected while the page requires a connection, go back to connect
  const requiresConn = !LOCAL_PAGES.includes(page)
  const effective: PageId = requiresConn && !connected ? 'connect' : page
  const Current = pages[effective]

  return (
    <Layout page={effective} onNavigate={navigate} onOpenPalette={() => setPaletteOpen(true)}>
      <Current />
      <Toasts />
      <CommandPalette open={paletteOpen} onClose={() => setPaletteOpen(false)} onNavigate={navigate} />
    </Layout>
  )
}

export default function App() {
  return (
    <StoreProvider>
      <Shell />
    </StoreProvider>
  )
}
