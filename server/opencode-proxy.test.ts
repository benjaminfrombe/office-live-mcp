import { mkdtempSync, writeFileSync } from 'node:fs'
import { createServer as createHttpServer, type Server } from 'node:http'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { handleOpenCodeProxy, OPENCODE_MOUNT_PATH, resetServiceCache } from './opencode-proxy.ts'

// Integration test: a mock OpenCode upstream + a real http server that routes
// through handleOpenCodeProxy, exercised with real fetch requests.

let upstream: Server
let upstreamPort: number
let proxy: Server
let proxyPort: number
let serviceFile: string
let receivedAuth: string | null = null
let receivedPath: string | null = null

beforeAll(async () => {
  // Mock OpenCode upstream: echoes auth header and path, supports streaming
  upstream = createHttpServer((req, res) => {
    receivedAuth = req.headers.authorization ?? null
    receivedPath = req.url ?? null
    if (req.url === '/api/event') {
      res.writeHead(200, { 'Content-Type': 'text/event-stream' })
      res.write('data: {"type":"session.text.delta","data":{"delta":"hi"}}\n\n')
      res.end()
      return
    }
    res.writeHead(200, { 'Content-Type': 'application/json' })
    res.end(JSON.stringify({ data: { echo: req.url } }))
  })
  await new Promise<void>((resolve) => upstream.listen(0, '127.0.0.1', resolve))
  upstreamPort = (upstream.address() as { port: number }).port

  serviceFile = join(mkdtempSync(join(tmpdir(), 'oc-proxy-test-')), 'service.json')
  writeFileSync(serviceFile, JSON.stringify({ url: `http://127.0.0.1:${upstreamPort}`, password: 'test-pw' }))

  proxy = createHttpServer((req, res) => {
    if (!handleOpenCodeProxy(req, res, req.url ?? '/')) {
      res.writeHead(404)
      res.end('not proxied')
    }
  })
  await new Promise<void>((resolve) => proxy.listen(0, '127.0.0.1', resolve))
  proxyPort = (proxy.address() as { port: number }).port
})

afterAll(() => {
  upstream.close()
  proxy.close()
})

describe('opencode bridge proxy', () => {
  it('does not claim URLs outside the mount path', () => {
    const calls: boolean[] = []
    calls.push(
      handleOpenCodeProxy(
        {} as Parameters<typeof handleOpenCodeProxy>[0],
        {} as Parameters<typeof handleOpenCodeProxy>[1],
        '/index.html',
      ),
    )
    expect(calls[0]).toBe(false)
  })

  it('forwards requests with injected basic auth and unwrapped path', async () => {
    resetServiceCache()
    process.env.OPENCODE_SERVICE_FILE = serviceFile
    const res = await fetch(`http://127.0.0.1:${proxyPort}${OPENCODE_MOUNT_PATH}/api/session/abc`)
    expect(res.status).toBe(200)
    const body = (await res.json()) as { data: { echo: string } }
    expect(body.data.echo).toBe('/api/session/abc')
    expect(receivedAuth).toBe(`Basic ${Buffer.from('opencode:test-pw').toString('base64')}`)
  })

  it('answers CORS preflight locally', async () => {
    const res = await fetch(`http://127.0.0.1:${proxyPort}${OPENCODE_MOUNT_PATH}/api/session`, {
      method: 'OPTIONS',
      headers: { Origin: 'https://localhost:55671' },
    })
    expect(res.status).toBe(204)
    expect(res.headers.get('access-control-allow-origin')).toBe('https://localhost:55671')
  })

  it('streams SSE through without buffering headers away', async () => {
    resetServiceCache()
    const res = await fetch(`http://127.0.0.1:${proxyPort}${OPENCODE_MOUNT_PATH}/api/event`)
    expect(res.status).toBe(200)
    const text = await res.text()
    expect(text).toContain('session.text.delta')
    expect(receivedPath).toBe('/api/event')
  })

  it('returns 503 when the service file is unreadable', async () => {
    process.env.OPENCODE_SERVICE_FILE = '/nonexistent/service.json'
    resetServiceCache()
    const res = await fetch(`http://127.0.0.1:${proxyPort}${OPENCODE_MOUNT_PATH}/api/session`)
    expect(res.status).toBe(503)
    expect(await res.json()).toHaveProperty('proxyError')
    // restore
    process.env.OPENCODE_SERVICE_FILE = serviceFile
    resetServiceCache()
  })
})
