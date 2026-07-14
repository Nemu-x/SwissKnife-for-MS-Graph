import { useCallback, useState } from 'react'
import { errMessage } from './api'

export interface AsyncState<T> {
  data: T | null
  error: string | null
  loading: boolean
  run: (fn: () => Promise<T>) => Promise<T | undefined>
  reset: () => void
  setData: (d: T | null) => void
}

// Imperative counterpart to TanStack Query for the button-to-request flow (ADR-003 amendment).
export function useAsync<T>(): AsyncState<T> {
  const [data, setData] = useState<T | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)

  const run = useCallback(async (fn: () => Promise<T>) => {
    setLoading(true)
    setError(null)
    try {
      const result = await fn()
      setData(result)
      return result
    } catch (e) {
      setError(errMessage(e))
      return undefined
    } finally {
      setLoading(false)
    }
  }, [])

  const reset = useCallback(() => {
    setData(null)
    setError(null)
    setLoading(false)
  }, [])

  return { data, error, loading, run, reset, setData }
}
