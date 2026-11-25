// Per-session token gating /opencode and /file (defense in depth on top of
// the Host-header check). The bridge issues a random token at startup and
// injects it into the served taskpane entry page; the add-in echoes it back
// as a `token` query param (img tags cannot send custom headers, so /file
// must accept it there). Local processes without the token are rejected —
// the Host check alone is a single control, this is the second lock.
//
// /mcp is deliberately NOT token-gated: external MCP clients connect
// out-of-band and cannot learn the taskpane token (see README "Security").

import { randomBytes, timingSafeEqual } from 'node:crypto'
import type { IncomingHttpHeaders } from 'node:http'

let token: string | null = null

/** Issue (or re-issue) the bridge token. Called once at bridge startup. */
export function issueToken(): string {
  token = randomBytes(32).toString('base64url')
  return token
}

/** The token issued for this bridge run, or null before startup. */
export function currentBridgeToken(): string | null {
  return token
}

/** Test hook — pin or clear the expected token. */
export function setBridgeToken(value: string | null): void {
  token = value
}

/** Constant-time comparison of two token strings (null-safe). */
function tokensEqual(a: string | null, b: string | null): boolean {
  if (!a || !b) return false
  const bufA = Buffer.from(a)
  const bufB = Buffer.from(b)
  if (bufA.length !== bufB.length) return false
  return timingSafeEqual(bufA, bufB)
}

/**
 * Check the `token` query param (or `x-bridge-token` header) of a request
 * against the issued token. False when no token was issued yet, when the
 * request carries none, or when it does not match.
 */
export function requestHasValidToken(req: { headers: IncomingHttpHeaders; url?: string }): boolean {
  if (!token) return false
  let provided: string | null = null
  const header = req.headers['x-bridge-token']
  if (typeof header === 'string' && header) provided = header
  if (!provided && req.url) {
    const query = new URL(req.url, 'http://localhost').searchParams.get('token')
    if (query) provided = query
  }
  return tokensEqual(provided, token)
}
