import { describe, expect, it } from 'vitest'
import { currentBridgeToken, issueToken, requestHasValidToken, setBridgeToken } from './bridge-token.ts'

function reqWith(url: string, header?: string): Parameters<typeof requestHasValidToken>[0] {
  return { headers: header ? { 'x-bridge-token': header } : {}, url }
}

describe('bridge token', () => {
  it('issues a random base64url token each call', () => {
    const a = issueToken()
    const b = issueToken()
    expect(a).toMatch(/^[A-Za-z0-9_-]{40,}$/)
    expect(a).not.toBe(b)
    expect(currentBridgeToken()).toBe(b)
  })

  it('accepts a matching token via query param', () => {
    setBridgeToken('tok123')
    expect(requestHasValidToken(reqWith('/file?path=/a.png&token=tok123'))).toBe(true)
    expect(requestHasValidToken(reqWith('/opencode/api/session?token=tok123'))).toBe(true)
  })

  it('accepts a matching token via header', () => {
    setBridgeToken('tok123')
    expect(requestHasValidToken(reqWith('/file?path=/a.png', 'tok123'))).toBe(true)
  })

  it('rejects a wrong, missing or empty token', () => {
    setBridgeToken('tok123')
    expect(requestHasValidToken(reqWith('/file?path=/a.png'))).toBe(false)
    expect(requestHasValidToken(reqWith('/file?path=/a.png&token=wrong'))).toBe(false)
    expect(requestHasValidToken(reqWith('/file?path=/a.png&token='))).toBe(false)
    expect(requestHasValidToken(reqWith('/file?path=/a.png', 'wrong'))).toBe(false)
  })

  it('rejects everything when no token was issued', () => {
    setBridgeToken(null)
    expect(requestHasValidToken(reqWith('/file?path=/a.png&token=tok123'))).toBe(false)
  })
})
