import { useState } from 'react'
import { Layout } from './components/Layout'
import { Toasts } from './components/Toasts'
import { StoreProvider, useStore } from './lib/store'
import { pages, type PageId } from './pages/registry'

function Shell() {
  const [page, setPage] = useState<PageId>('connect')
  const { connected } = useStore()

  // if disconnected while the page requires a connection, go back to connect
  const requiresConn = page !== 'connect' && page !== 'settings'
  const effective: PageId = requiresConn && !connected ? 'connect' : page
  const Current = pages[effective]

  return (
    <Layout page={effective} onNavigate={setPage}>
      <Current />
      <Toasts />
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
