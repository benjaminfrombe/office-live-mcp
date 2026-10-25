import { createServer, type Server } from 'node:http'
import { PassThrough } from 'node:stream'
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { type EnsureDaemonResult, ensureDaemon, parseSseMessages, runStdioProxy, stopDaemon } from './stdio-proxy.ts'

// Each fake daemon binds an ephemeral port (listen(0)) and points the module
// under test at it via MCP_PORT — no cross-test or cross-run port collisions,
// and no undici keep-alive races against servers that are already gone.
function startFakeDaemon(handlers: {
  health?: () => { status: number; body: unknown }
  mcpPost?: (
    body: string,
    req: { headers: Record<string, string | string[] | undefined> },
  ) => {
    status: number
    contentType: string
    body: string
    sessionId?: string
  }
  shutdown?: () => boolean
}): Promise<Server> {
  return new Promise((resolve, reject) => {
    const server = createServer((req, res) => {
      const url = (req.url ?? '').split('?')[0]
      let body = ''
      req.on('data', (c: Buffer) => {
        body += c.toString()
      })
      req.on('end', () => {
        if (url === '/health' && handlers.health) {
          const h = handlers.health()
          res.writeHead(h.status, { 'Content-Type': 'application/json' })
          res.end(JSON.stringify(h.body))
        } else if (url === '/mcp' && handlers.mcpPost) {
          const out = handlers.mcpPost(body, req)
          const headers: Record<string, string> = { 'Content-Type': out.contentType }
          if (out.sessionId) headers['Mcp-Session-Id'] = out.sessionId
          res.writeHead(out.status, headers)
          res.end(out.body)
        } else if (url === '/shutdown' && handlers.shutdown) {
          const stopped = handlers.shutdown()
          res.writeHead(stopped ? 200 : 404)
          res.end()
        } else {
          res.writeHead(404)
          res.end()
        }
      })
    })
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      if (address === null || typeof address === 'string') {
        reject(new Error('Failed to determine the ephemeral test port'))
        return
      }
      process.env.MCP_PORT = String(address.port)
      resolve(server)
    })
  })
}

function closeServer(server: Server): Promise<void> {
  // Node's fetch (undici) keeps keep-alive sockets open for a few seconds,
  // which would leave the port bound and break the next test's listen().
  server.closeAllConnections?.()
  return new Promise<void>((r) => server.close(() => r()))
}

describe('parseSseMessages', () => {
  it('extracts JSON-RPC messages from SSE data lines', () => {
    const body = [
      'event: message',
      'data: {"jsonrpc":"2.0","id":1,"result":{"ok":true}}',
      '',
      'event: message',
      'data: {"jsonrpc":"2.0","method":"notifications/tools_list_changed"}',
      '',
      '',
    ].join('\n')
    const messages = parseSseMessages(body)
    expect(messages).toHaveLength(2)
    expect((messages[0] as { id: number }).id).toBe(1)
    expect((messages[1] as { method: string }).method).toBe('notifications/tools_list_changed')
  })

  it('ignores [DONE] sentinels and malformed lines', () => {
    const messages = parseSseMessages('data: [DONE]\ndata: not-json\n: keepalive comment\n')
    expect(messages).toEqual([])
  })
})

describe('ensureDaemon', () => {
  it('returns alreadyRunning when a healthy daemon answers', async () => {
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', version: '0.6.0', bridge: true } }),
    })
    try {
      const result = await ensureDaemon('0.6.0')
      expect(result.alreadyRunning).toBe(true)
      expect(result.health.version).toBe('0.6.0')
    } finally {
      await closeServer(server)
    }
  })

  it('throws when the daemon never becomes healthy', async () => {
    // Nothing listens here (port 1 cannot be bound by userland): the spawn
    // seam is a no-op, so the start timeout is the only way out.
    process.env.MCP_PORT = '1'
    const spawn = vi.fn(() => -1)
    await expect(ensureDaemon('0.6.0', { spawnDaemon: spawn, startTimeoutMs: 300 })).rejects.toThrow(
      /did not become healthy/,
    )
    expect(spawn).toHaveBeenCalledTimes(1)
  })

  it('starts a daemon when none is running and health appears', async () => {
    // Health only answers after the spawn seam has been called.
    let spawned = false
    const server = await startFakeDaemon({
      health: () =>
        spawned
          ? { status: 200, body: { status: 'ok', version: '0.6.0', bridge: true } }
          : { status: 503, body: { status: 'starting' } },
    })
    try {
      const spawn = vi.fn(() => {
        spawned = true
        return 12345
      })
      const result = await ensureDaemon('0.6.0', { spawnDaemon: spawn })
      expect(result.alreadyRunning).toBe(false)
      expect(result.health.version).toBe('0.6.0')
      expect(spawn).toHaveBeenCalledTimes(1)
    } finally {
      await closeServer(server)
    }
  })
})

describe('runStdioProxy', () => {
  it('proxies initialize, captures the session id, and forwards follow-up calls', async () => {
    let capturedSessionId: string | undefined
    const server = await startFakeDaemon({
      mcpPost: (body, req) => {
        const parsed = JSON.parse(body) as { method?: string; id?: number }
        if (parsed.method === 'initialize') {
          return {
            status: 200,
            contentType: 'text/event-stream',
            sessionId: 'sess-123',
            body: `event: message\ndata: ${JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { serverInfo: { name: 'office-live-mcp' } } })}\n\n`,
          }
        }
        if (parsed.id === undefined) {
          // Notification: 202 Accepted, no body — like the real daemon.
          return { status: 202, contentType: 'text/plain', body: '' }
        }
        capturedSessionId = req.headers['mcp-session-id'] as string | undefined
        return {
          status: 200,
          contentType: 'application/json',
          body: JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { tools: [] } }),
        }
      },
    })

    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))

    const ensure = vi.fn<() => Promise<EnsureDaemonResult>>(() =>
      Promise.resolve({ alreadyRunning: true, health: { status: 'ok', version: '0.6.0' } }),
    )

    const done = runStdioProxy({ stdin, stdout, localVersion: '0.6.0', ensure })

    // All lines in a single chunk: follow-ups must queue behind the
    // initialize round-trip that carries the session id.
    stdin.write(
      [
        JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} }),
        JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }),
        JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }),
      ].join('\n') + '\n',
    )
    await new Promise((r) => setTimeout(r, 200))
    stdin.end()
    await done

    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(2)
    expect((lines[0] as { id: number }).id).toBe(1)
    expect((lines[1] as { id: number }).id).toBe(2)
    // The proxy must have sent the captured session id on the second request.
    expect(capturedSessionId).toBe('sess-123')
    expect(ensure).toHaveBeenCalledTimes(1)

    await closeServer(server)
  })

  it('emits a JSON-RPC internal error when the daemon is unreachable', async () => {
    process.env.MCP_PORT = '1'
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))

    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      ensure: () => Promise.resolve({ alreadyRunning: true, health: { status: 'ok' } }),
    })

    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 7, method: 'tools/list', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 150))
    stdin.end()
    await done

    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(1)
    expect((lines[0] as { error: { code: number } }).error.code).toBe(-32603)
  })
})

describe('stopDaemon', () => {
  it('reports whether a daemon was stopped', async () => {
    let shutdownHit = false
    const server = await startFakeDaemon({
      shutdown: () => {
        shutdownHit = true
        return true
      },
    })
    try {
      const stopped = await stopDaemon()
      expect(stopped).toBe(true)
      expect(shutdownHit).toBe(true)
    } finally {
      await closeServer(server)
    }
  })

  it('returns false when nothing is running', async () => {
    process.env.MCP_PORT = '1'
    const stopped = await stopDaemon()
    expect(stopped).toBe(false)
  })
})

afterAll(() => {
  vi.restoreAllMocks()
})
