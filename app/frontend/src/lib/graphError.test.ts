import { describe, it, expect } from 'vitest'
import { parseErr } from './graphError'

describe('parseErr', () => {
  it('parses the operr envelope', () => {
    const p = parseErr('operr:{"code":"ErrorAccessDenied","status":403,"requestId":"r-1","hint":"Mail.ReadWrite","message":"Access is denied."}')
    expect(p).toEqual({
      message: 'Access is denied.',
      code: 'ErrorAccessDenied',
      status: 403,
      requestId: 'r-1',
      hint: 'Mail.ReadWrite',
    })
  })

  it('parses the raw GraphError string with requestId', () => {
    const p = parseErr('graph: 403 ErrorAccessDenied: Access is denied. Check credentials and try again. (requestId=9d4e3c6f)')
    expect(p.status).toBe(403)
    expect(p.code).toBe('ErrorAccessDenied')
    expect(p.requestId).toBe('9d4e3c6f')
    expect(p.message).toBe('Access is denied. Check credentials and try again.')
  })

  it('parses the raw GraphError string without requestId', () => {
    const p = parseErr('graph: 404 itemNotFound: The resource could not be found.')
    expect(p.status).toBe(404)
    expect(p.code).toBe('itemNotFound')
    expect(p.requestId).toBeUndefined()
  })

  it('passes plain strings through', () => {
    expect(parseErr('not connected — connect to a tenant first')).toEqual({
      message: 'not connected — connect to a tenant first',
    })
  })

  it('falls back to plain on malformed operr payloads', () => {
    const raw = 'operr:{broken json'
    expect(parseErr(raw).message).toBe(raw)
  })
})
