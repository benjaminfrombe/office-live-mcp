// Host-header validation for the loopback-only servers (bridge :8791, MCP HTTP
// :8792). Both bind 127.0.0.1, but a DNS-rebinding page can still reach them:
// the browser resolves the attacker's hostname to 127.0.0.1 and sends the
// attacker's hostname in the Host header. Rejecting any Host that is not this
// machine closes that route (the same protection the MCP SDK applies to /mcp
// via enableDnsRebindingProtection, extended to the routes it does not cover:
// /opencode, /file, /shutdown and the static assets).
//
// Allowed: localhost, 127.0.0.0/8, ::1, the machine's hostname and its own
// interface addresses (the bridge can bind all interfaces via BRIDGE_HOST for
// an Office VM on the LAN, and the taskpane may then be served via any of
// this host's own addresses — same set as the WebSocket origin allowlist).

import { hostname, networkInterfaces } from 'node:os'

let allowedHosts: Set<string> | null = null

function allowedHostSet(): Set<string> {
  if (allowedHosts) return allowedHosts
  const set = new Set<string>(['localhost', '127.0.0.1', '::1', '[::1]', hostname().toLowerCase()])
  for (const ifaces of Object.values(networkInterfaces())) {
    for (const iface of ifaces ?? []) {
      const addr = iface.address
      set.add(addr)
      if (iface.family === 'IPv6') set.add(`[${addr}]`)
    }
  }
  allowedHosts = set
  return set
}

/** Test hook — recompute the allowed-host set (hostname/interfaces can differ per test env). */
export function resetHostCheckCache(): void {
  allowedHosts = null
}

/**
 * Validate the request's Host header against this machine's own names and
 * addresses. Returns false for a missing Host (not a browser or HTTP/1.1
 * client) and for anything pointing at another name — the DNS-rebinding case.
 */
export function hasLocalHostHeader(req: { headers: { host?: string | string[] } }): boolean {
  const raw = req.headers.host
  if (typeof raw !== 'string' || raw.length === 0) return false
  // "Host: host:port" — strip the port; IPv6 literals arrive bracketed and
  // the port follows the closing bracket.
  const host = raw.startsWith('[') ? raw.slice(0, raw.indexOf(']') + 1) : (raw.split(':')[0] ?? raw)
  return allowedHostSet().has(host.toLowerCase())
}
