/**
 * OpenCode reverse proxy built into the bridge server.
 *
 * The taskpane needs to reach the OpenCode HTTP API, but OpenCode always
 * requires basic auth with a per-service password and never reflects CORS for
 * https origins. Serving the proxy from the bridge itself means the add-in
 * talks to its own origin (`https://localhost:<bridge>/opencode/...`) — no
 * CORS, no extra certificate, no separate process.
 *
 * The OpenCode service URL + password are read from the background service
 * registration file and cached for 2s, so a service restart (new port, new
 * password) is picked up automatically.
 */
import { readFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { request as httpRequest } from 'node:http'
import { homedir } from 'node:os'
import { join } from 'node:path'

const MOUNT_PATH = '/opencode'

export const OPENCODE_MOUNT_PATH = MOUNT_PATH

interface ServiceInfo {
  url: string
  password: string
}

let cached: { target: URL; password: string } | null = null
let lastRead = 0

function serviceFilePath(): string {
  // Read per call so tests (and runtime overrides) can point at another file
  return process.env.OPENCODE_SERVICE_FILE || join(homedir(), '.local/state/opencode/service.json')
}

/** Read (and briefly cache) the current OpenCode service endpoint. */
export function service(): { target: URL; password: string } {
  if (cached && Date.now() - lastRead < 2000) return cached
  const info = JSON.parse(readFileSync(serviceFilePath(), 'utf8')) as ServiceInfo
  cached = { target: new URL(info.url), password: info.password }
  lastRead = Date.now()
  return cached
}

/** Test hook — reset the cached service endpoint. */
export function resetServiceCache(): void {
  cached = null
  lastRead = 0
}

function corsHeaders(req: IncomingMessage): Record<string, string> {
  const origin = req.headers.origin
  if (!origin) return {}
  return {
    'access-control-allow-origin': origin,
    'access-control-allow-credentials': 'true',
  }
}

function pipe(req: IncomingMessage, res: ServerResponse, target: URL, password: string, forwardPath: string): void {
  const upstream = httpRequest(
    {
      protocol: target.protocol,
      hostname: target.hostname,
      port: target.port,
      path: forwardPath,
      method: req.method,
      headers: {
        ...req.headers,
        host: `${target.hostname}:${target.port}`,
        authorization: `Basic ${Buffer.from(`opencode:${password}`).toString('base64')}`,
      },
    },
    (upstreamRes) => {
      const headers = { ...upstreamRes.headers, ...corsHeaders(req) }
      res.writeHead(upstreamRes.statusCode || 502, headers)
      // Stream the body — SSE event streams must flow chunk by chunk.
      upstreamRes.pipe(res)
    },
  )
  upstream.on('error', (err) => {
    if (!res.headersSent) {
      res.writeHead(502, { 'Content-Type': 'application/json', ...corsHeaders(req) })
    }
    res.end(JSON.stringify({ proxyError: String(err.message) }))
  })
  req.pipe(upstream)
}

/**
 * Handle `/opencode/...` requests. Returns false when the URL is not on the
 * mount path so the caller can fall through to other routes.
 */
export function handleOpenCodeProxy(req: IncomingMessage, res: ServerResponse, rawUrl: string): boolean {
  const [pathOnly] = rawUrl.split('?')
  if (pathOnly !== MOUNT_PATH && !pathOnly.startsWith(`${MOUNT_PATH}/`)) {
    return false
  }

  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'access-control-allow-origin': req.headers.origin || '*',
      'access-control-allow-methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
      'access-control-allow-headers': 'authorization, content-type',
      'access-control-max-age': '86400',
    })
    res.end()
    return true
  }

  // Forward the full path including query string
  const forwardPath = rawUrl.slice(MOUNT_PATH.length) || '/'
  try {
    const { target, password } = service()
    pipe(req, res, target, password, forwardPath)
  } catch (err) {
    res.writeHead(503, { 'Content-Type': 'application/json', ...corsHeaders(req) })
    res.end(JSON.stringify({ proxyError: `OpenCode service file unreadable: ${String(err)}` }))
  }
  return true
}
