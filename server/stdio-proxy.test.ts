// Isolate the daemon dir BEFORE importing the module under test: it reads
// OLM_DAEMON_DIR at module load, and the real ~/.office-live-mcp/daemon.lock
// (with a live PID when a daemon is running) would change the code paths.
vi.hoisted(() => {
  process.env.OLM_DAEMON_DIR = '/tmp/olm-stdio-proxy-test'
  process.env.OLM_DAEMON_LOG = '/tmp/olm-stdio-proxy-test/daemon.log'
})

import { mkdirSync, readFileSync, rmSync, unlinkSync, writeFileSync } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { PassThrough } from 'node:stream'
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { type EnsureDaemonResult, ensureDaemon, parseSseMessages, runStdioProxy, stopDaemon } from './stdio-proxy.ts'

const TEST_DAEMON_DIR = '/tmp/olm-stdio-proxy-test'
const TEST_LOCK_FILE = `${TEST_DAEMON_DIR}/daemon.lock`

beforeEach(() => {
  // Fresh daemon dir per test so lock-file state never leaks between tests
  // (or from a real running daemon).
  rmSync(TEST_DAEMON_DIR, { recursive: true, force: true })
  mkdirSync(TEST_DAEMON_DIR, { recursive: true })
})

afterAll(() => {
  rmSync(TEST_DAEMON_DIR, { recursive: true, force: true })
})

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

  it('restarts a stale daemon when versions mismatch', async () => {
    // The fake daemon reports the old version until it receives /shutdown,
    // after which it reports the new one (simulating a fresh daemon on the
    // same endpoint).
    let version = '0.5.0'
    let shutdownHit = false
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', version, bridge: true } }),
      shutdown: () => {
        shutdownHit = true
        version = '0.6.0'
        return true
      },
    })
    try {
      const spawn = vi.fn(() => 12345)
      const result = await ensureDaemon('0.6.0', { spawnDaemon: spawn, shutdownWaitMs: 200 })
      expect(shutdownHit).toBe(true)
      expect(spawn).toHaveBeenCalledTimes(1)
      expect(result.alreadyRunning).toBe(false)
      expect(result.health.version).toBe('0.6.0')
    } finally {
      await closeServer(server)
    }
  })

  it('does not restart a daemon whose version matches', async () => {
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', version: '0.6.0', bridge: true } }),
    })
    try {
      const stopFn = vi.fn(async () => true)
      const spawn = vi.fn(() => -1)
      const result = await ensureDaemon('0.6.0', { spawnDaemon: spawn, stopDaemonFn: stopFn })
      expect(result.alreadyRunning).toBe(true)
      expect(stopFn).not.toHaveBeenCalled()
      expect(spawn).not.toHaveBeenCalled()
    } finally {
      await closeServer(server)
    }
  })

  it('treats a daemon without a version field as matching and leaves it running', async () => {
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', bridge: true } }),
    })
    try {
      const stopFn = vi.fn(async () => true)
      const result = await ensureDaemon('0.6.0', { spawnDaemon: () => -1, stopDaemonFn: stopFn })
      expect(result.alreadyRunning).toBe(true)
      expect(stopFn).not.toHaveBeenCalled()
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

  it('heals a stale or missing daemon.lock from the health-reported pid', async () => {
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', version: '0.6.0', bridge: true, pid: 424242 } }),
    })
    const result = await ensureDaemon('0.6.0')
    expect(result.alreadyRunning).toBe(true)
    expect(JSON.parse(readFileSync(TEST_LOCK_FILE, 'utf8'))).toMatchObject({ pid: 424242 })
    await closeServer(server)
  })

  it('does not clobber a daemon.lock held by a live process', async () => {
    writeFileSync(TEST_LOCK_FILE, JSON.stringify({ pid: process.pid, startedAt: 'x' }))
    const server = await startFakeDaemon({
      health: () => ({ status: 200, body: { status: 'ok', version: '0.6.0', bridge: true, pid: 424242 } }),
    })
    await ensureDaemon('0.6.0')
    expect(JSON.parse(readFileSync(TEST_LOCK_FILE, 'utf8')).pid).toBe(process.pid)
    await closeServer(server)
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

  it('skips spawning when a live process holds the lock, and still waits for health', async () => {
    // Simulate another wrapper that just spawned the daemon: lock exists with
    // a live PID, health appears a moment later.
    writeFileSync(TEST_LOCK_FILE, JSON.stringify({ pid: process.pid }))
    let healthy = false
    const server = await startFakeDaemon({
      health: () =>
        healthy
          ? { status: 200, body: { status: 'ok', version: '0.6.0', bridge: true } }
          : { status: 503, body: { status: 'starting' } },
    })
    try {
      const spawn = vi.fn(() => {
        throw new Error('must not be called when the lock is held by a live process')
      })
      const result = ensureDaemon('0.6.0', { spawnDaemon: spawn }).then((r) => {
        return r
      })
      // Health comes up after the check started.
      setTimeout(() => {
        healthy = true
      }, 100)
      const out = await result
      expect(out.alreadyRunning).toBe(false)
      expect(out.health.version).toBe('0.6.0')
      expect(spawn).not.toHaveBeenCalled()
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
      Promise.resolve({
        alreadyRunning: true,
        health: { status: 'ok', version: '0.6.0' },
        base: `http://127.0.0.1:${process.env.MCP_PORT}`,
      }),
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
      ensure: () => Promise.resolve({ alreadyRunning: true, health: { status: 'ok' }, base: 'http://127.0.0.1:1' }),
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
  it('does not let a slow request block later requests once the session is initialized', async () => {
    const server = await new Promise<Server>((resolve) => {
      const srv = createServer((req, res) => {
        let body = ''
        req.on('data', (c: Buffer) => {
          body += c.toString()
        })
        req.on('end', () => {
          const parsed = JSON.parse(body) as { method?: string; id?: number }
          const reply = (delayMs: number, headers: Record<string, string> = {}) =>
            setTimeout(() => {
              res.writeHead(200, { 'Content-Type': 'application/json', ...headers })
              res.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }))
            }, delayMs)
          if (parsed.method === 'initialize') reply(0, { 'Mcp-Session-Id': 's1' })
          else if (parsed.method === 'tools/call') reply(600)
          else reply(0)
        })
      })
      srv.listen(0, '127.0.0.1', () => {
        process.env.MCP_PORT = String((srv.address() as { port: number }).port)
        resolve(srv)
      })
    })
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const arrival: number[] = []
    stdout.on('data', (c: Buffer) => {
      for (const l of c.toString().trim().split('\n')) arrival.push((JSON.parse(l) as { id: number }).id)
    })
    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      ensure: () =>
        Promise.resolve({
          alreadyRunning: true,
          health: { status: 'ok' },
          base: `http://127.0.0.1:${process.env.MCP_PORT}`,
        }),
    })
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 100))
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: {} })}\n`)
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 3, method: 'tools/list', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 200))
    // The fast request answers while the slow one is still pending.
    expect(arrival).toEqual([1, 3])
    stdin.end()
    await done
    // end waits for in-flight requests: the slow one is answered too.
    expect(arrival).toEqual([1, 3, 2])
    await closeServer(server)
  })

  it('answers a failed request with that request id so the client does not wait for a timeout', async () => {
    process.env.MCP_PORT = '1'
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))
    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      ensure: () => Promise.resolve({ alreadyRunning: true, health: { status: 'ok' }, base: 'http://127.0.0.1:1' }),
    })
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 7, method: 'tools/call', params: {} })}\n`)
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' })}\n`)
    await new Promise((r) => setTimeout(r, 150))
    stdin.end()
    await done
    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l) as { id: unknown })
    // One error for the request (carrying its id); none for the notification.
    expect(lines).toHaveLength(1)
    expect(lines[0].id).toBe(7)
  })

  it('re-discovers the daemon and retries when the daemon dies mid-session', async () => {
    // Daemon A answers initialize, then dies before the next request — the
    // situation after a version-skew daemon restart while proxies stay up.
    const serverA = await startFakeDaemon({
      mcpPost: () => ({
        status: 200,
        contentType: 'application/json',
        sessionId: 's1',
        body: JSON.stringify({ jsonrpc: '2.0', id: 1, result: { initialized: true } }),
      }),
    })
    const portA = process.env.MCP_PORT
    let serverBSeenSession: string | undefined
    const serverB = await startFakeDaemon({
      mcpPost: (body, req) => {
        const sid = req.headers['mcp-session-id']
        serverBSeenSession = Array.isArray(sid) ? sid[0] : sid
        return {
          status: 200,
          contentType: 'application/json',
          body: JSON.stringify({ jsonrpc: '2.0', id: JSON.parse(body).id, result: { ok: true } }),
        }
      },
    })
    const portB = process.env.MCP_PORT

    let ensureCalls = 0
    const bases = [portA, portB]
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))
    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      ensure: () =>
        Promise.resolve({
          alreadyRunning: ensureCalls > 0,
          health: { status: 'ok' },
          base: `http://127.0.0.1:${bases[Math.min(ensureCalls++, bases.length - 1)]}`,
        }),
    })

    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 150))
    await closeServer(serverA)
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 300))
    stdin.end()
    await done
    void serverB

    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    // Both requests answered — the second one after re-discovery, not with an error.
    expect(lines.map((l) => l.id)).toEqual([1, 2])
    expect((lines[1] as { result: unknown }).result).toEqual({ ok: true })
    expect(ensureCalls).toBe(2)
    // The stale session id must not leak to the fresh daemon.
    expect(serverBSeenSession).toBeUndefined()
    await closeServer(serverB)
  })

  it('recovers the session when the daemon reports it as invalid', async () => {
    // A restarted daemon rejects the proxy's old session id with -32000; the
    // proxy must drop the session, re-discover, and retry instead of passing
    // the daemon's error through to the client.
    let rejectedOnce = false
    let retriedWithoutSession = false
    const server = await startFakeDaemon({
      mcpPost: (body, req) => {
        const parsed = JSON.parse(body) as { id?: number; method?: string }
        if (parsed.method === 'initialize') {
          return {
            status: 200,
            contentType: 'application/json',
            sessionId: 'old-sess',
            body: JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { initialized: true } }),
          }
        }
        if (req.headers['mcp-session-id'] === 'old-sess' && !rejectedOnce) {
          rejectedOnce = true
          return {
            status: 400,
            contentType: 'application/json',
            body: JSON.stringify({
              jsonrpc: '2.0',
              error: { code: -32000, message: 'Bad request: no valid session' },
              id: parsed.id ?? null,
            }),
          }
        }
        retriedWithoutSession = !req.headers['mcp-session-id']
        return {
          status: 200,
          contentType: 'application/json',
          body: JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { ok: true } }),
        }
      },
    })

    let ensureCalls = 0
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))
    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      ensure: () => {
        ensureCalls++
        return Promise.resolve({
          alreadyRunning: ensureCalls > 1,
          health: { status: 'ok' },
          base: `http://127.0.0.1:${process.env.MCP_PORT}`,
        })
      },
    })

    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 150))
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 300))
    stdin.end()
    await done

    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    // The client sees the recovered result, not the daemon's -32000.
    expect(lines.map((l) => l.id)).toEqual([1, 2])
    expect((lines[1] as { result: unknown }).result).toEqual({ ok: true })
    expect(ensureCalls).toBe(2)
    expect(retriedWithoutSession).toBe(true)
    await closeServer(server)
  })

  it('answers with an error instead of blocking forever when the daemon hangs', async () => {
    // Raw server that accepts the request but never responds — a wedged
    // keep-alive socket after a daemon kill must not block the proxy forever.
    const server = await new Promise<Server>((resolve) => {
      const srv = createServer(() => undefined) // no response ever
      srv.listen(0, '127.0.0.1', () => {
        process.env.MCP_PORT = String((srv.address() as { port: number }).port)
        resolve(srv)
      })
    })
    const stdin = new PassThrough()
    const stdout = new PassThrough()
    const chunks: string[] = []
    stdout.on('data', (c: Buffer) => chunks.push(c.toString()))
    const done = runStdioProxy({
      stdin,
      stdout,
      localVersion: '0.6.0',
      requestTimeoutMs: 100,
      ensure: () =>
        Promise.resolve({
          alreadyRunning: true,
          health: { status: 'ok' },
          base: `http://127.0.0.1:${process.env.MCP_PORT}`,
        }),
    })
    stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: 9, method: 'tools/call', params: {} })}\n`)
    await new Promise((r) => setTimeout(r, 500))
    stdin.end()
    await done
    const lines = chunks
      .join('')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(1)
    expect((lines[0] as { error: { code: number } }).error.code).toBe(-32603)
    await closeServer(server)
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
