/**
 * Local passwordless proxy in front of the OpenCode HTTP server.
 *
 * OpenCode's HTTP API always requires basic auth with a per-service password
 * (no flag to disable it). The Office add-in taskpane should not need to know
 * that password, so this proxy listens on a fixed port, injects the
 * `Authorization` header, and re-reads the password whenever the service
 * restarts (new port + new password in service.json).
 *
 * The add-in configures `http://127.0.0.1:4096` with an empty password.
 *
 * Run: node --experimental-strip-types ./scripts/opencode-proxy.ts
 * Env: OPENCODE_PROXY_PORT (default 4096), OPENCODE_SERVICE_FILE
 */
import { readFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { createServer, request as httpRequest } from 'node:http'
import type { ServerOptions } from 'node:https'
import { createServer as createHttpsServer } from 'node:https'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const PORT = Number(process.env.OPENCODE_PROXY_PORT || 4096)
const SERVICE_FILE = process.env.OPENCODE_SERVICE_FILE || join(homedir(), '.local/state/opencode/service.json')
// Office's WKWebView blocks mixed content: an add-in page served over HTTPS
// cannot fetch http:// URLs. Set OPENCODE_PROXY_TLS=1 to serve the proxy over
// TLS with the same mkcert dev certificates the bridge uses, and configure the
// add-in with https://127.0.0.1:4096 instead.
const CERT_DIR = process.env.OPENCODE_PROXY_CERT_DIR || join(fileURLToPath(new URL('.', import.meta.url)), '../certs')
const useTls = process.env.OPENCODE_PROXY_TLS === '1'

interface ServiceInfo {
  url: string
  password: string
}

let cached: { target: URL; password: string } | null = null
let lastRead = 0

function service(): { target: URL; password: string } {
  // Cache for 2s so bursts of requests don't hammer the file, but a service
  // restart is picked up almost immediately.
  if (cached && Date.now() - lastRead < 2000) return cached
  const info = JSON.parse(readFileSync(SERVICE_FILE, 'utf8')) as ServiceInfo
  cached = { target: new URL(info.url), password: info.password }
  lastRead = Date.now()
  return cached
}

function pipe(req: IncomingMessage, res: ServerResponse, target: URL, password: string) {
  const upstream = httpRequest(
    {
      protocol: target.protocol,
      hostname: target.hostname,
      port: target.port,
      path: req.url,
      method: req.method,
      headers: {
        ...req.headers,
        host: `${target.hostname}:${target.port}`,
        // Inject auth — the client never sees the password.
        authorization: `Basic ${Buffer.from(`opencode:${password}`).toString('base64')}`,
      },
    },
    (upstreamRes) => {
      const headers = { ...upstreamRes.headers }
      // OpenCode only reflects CORS for http origins, never https — but the
      // taskpane is served over https. The proxy owns CORS instead: echo the
      // caller's origin so the browser accepts the response.
      if (req.headers.origin) {
        headers['access-control-allow-origin'] = req.headers.origin
        headers['access-control-allow-credentials'] = 'true'
      }
      res.writeHead(upstreamRes.statusCode || 502, headers)
      // Stream the body — SSE event streams must flow chunk by chunk.
      upstreamRes.pipe(res)
    },
  )
  upstream.on('error', (err) => {
    if (!res.headersSent) {
      res.writeHead(502, { 'Content-Type': 'application/json' })
    }
    res.end(JSON.stringify({ proxyError: String(err.message) }))
  })
  req.pipe(upstream)
}

const handler = (req: IncomingMessage, res: ServerResponse) => {
  // CORS preflight — answer here, never forwarded upstream.
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'access-control-allow-origin': req.headers.origin || '*',
      'access-control-allow-methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
      'access-control-allow-headers': 'authorization, content-type',
      'access-control-max-age': '86400',
    })
    res.end()
    return
  }
  try {
    const { target, password } = service()
    pipe(req, res, target, password)
  } catch (err) {
    res.writeHead(503, { 'Content-Type': 'application/json' })
    res.end(JSON.stringify({ proxyError: `OpenCode service file unreadable: ${String(err)}` }))
  }
}

const server = useTls
  ? createHttpsServer(
      {
        key: readFileSync(join(CERT_DIR, 'localhost-key.pem')),
        cert: readFileSync(join(CERT_DIR, 'localhost.pem')),
      } satisfies ServerOptions,
      handler,
    )
  : createServer(handler)

server.listen(PORT, '127.0.0.1', () => {
  console.log(`opencode-proxy listening on ${useTls ? 'https' : 'http'}://127.0.0.1:${PORT} -> ${SERVICE_FILE}`)
})

// Graceful shutdown
process.on('SIGINT', () => server.close(() => process.exit(0)))
process.on('SIGTERM', () => server.close(() => process.exit(0)))
