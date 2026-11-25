import { hostname, networkInterfaces } from 'node:os'
import { describe, expect, it } from 'vitest'
import { hasLocalHostHeader, resetHostCheckCache } from './host-check.ts'

function reqWith(host: string | undefined): Parameters<typeof hasLocalHostHeader>[0] {
  return { headers: host === undefined ? {} : { host } }
}

describe('hasLocalHostHeader', () => {
  it("accepts this machine's own names and addresses", () => {
    resetHostCheckCache()
    expect(hasLocalHostHeader(reqWith('localhost:8791'))).toBe(true)
    expect(hasLocalHostHeader(reqWith('127.0.0.1:8791'))).toBe(true)
    expect(hasLocalHostHeader(reqWith('LOCALHOST:8791'))).toBe(true)
    expect(hasLocalHostHeader(reqWith(`${hostname().toLowerCase()}:8791`))).toBe(true)
  })

  it('accepts IPv6 loopback with brackets', () => {
    expect(hasLocalHostHeader(reqWith('[::1]:8791'))).toBe(true)
  })

  it("accepts this host's own LAN address (BRIDGE_HOST=0.0.0.0 serving)", () => {
    const lan = Object.values(networkInterfaces())
      .flat()
      .find((iface) => iface && !iface.internal && iface.family === 'IPv4')?.address
    if (lan) expect(hasLocalHostHeader(reqWith(`${lan}:8791`))).toBe(true)
  })

  it('rejects a foreign hostname (DNS rebinding)', () => {
    expect(hasLocalHostHeader(reqWith('evil.example:8791'))).toBe(false)
    expect(hasLocalHostHeader(reqWith('evil.example'))).toBe(false)
  })

  it('rejects a missing or empty Host header', () => {
    expect(hasLocalHostHeader(reqWith(undefined))).toBe(false)
    expect(hasLocalHostHeader(reqWith(''))).toBe(false)
  })
})
