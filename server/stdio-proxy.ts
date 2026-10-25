// STDIO → HTTP daemon proxy.
//
// When a harness (Claude Code, OpenCode, Claude Desktop, ...) loads this MCP
// server over stdio, this module makes sure a single shared daemon is running
// (`--http --bridge`, detached) and then proxies newline-delimited JSON-RPC
// from stdin/stdout to the daemon's HTTP MCP endpoint.
//
// Why a proxy instead of running the bridge in-process: the bridge is a
// single-instance server (one process binds :8080/:8443/:3001) that must stay
// alive for the Office add-in even when no harness session is open. Spawning a
// bridge per harness session means port conflicts (EADDRINUSE) and a bridge
// that dies with the session. With the proxy, the first harness to load the
// MCP server starts the daemon detached; every harness (and the add-in) shares
// that one instance, and the daemon survives harness shutdowns.

import { spawn } from 'node:child_process'
import { closeSync, mkdirSync, openSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const HEALTH_TIMEOUT_MS = 1_000
const DAEMON_START_TIMEOUT_MS = 15_000
const DAEMON_POLL_INTERVAL_MS = 200

/** Base URL of the daemon's HTTP MCP server (read per call so tests can point it elsewhere). */
function mcpHttpBase(): string {
  return `http://127.0.0.1:${Number(process.env.MCP_PORT) || 3001}`
}

const DAEMON_DIR = process.env.OLM_DAEMON_DIR || join(homedir(), '.office-live-mcp')
const DAEMON_LOCK_FILE = join(DAEMON_DIR, 'daemon.lock')
const DAEMON_LOG_FILE = process.env.OLM_DAEMON_LOG || join(DAEMON_DIR, 'daemon.log')

// ---------------------------------------------------------------------------
// Health check
// ---------------------------------------------------------------------------

export interface DaemonHealth {
  status: string
  version?: string
  bridge?: boolean
  [key: string]: unknown
}

/** Probe the daemon's /health endpoint. Returns undefined when unreachable. */
export async function fetchDaemonHealth(timeoutMs = HEALTH_TIMEOUT_MS): Promise<DaemonHealth | undefined> {
  try {
    const res = await fetch(`${mcpHttpBase()}/health`, { signal: AbortSignal.timeout(timeoutMs) })
    if (!res.ok) return undefined
    return (await res.json()) as DaemonHealth
  } catch {
    return undefined
  }
}

// ---------------------------------------------------------------------------
// Lock file (records the daemon PID so `--stop-daemon` can find it)
// ---------------------------------------------------------------------------

function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (err: unknown) {
    return !(typeof err === 'object' && err !== null && 'code' in err && (err as { code?: string }).code === 'ESRCH')
  }
}

export function readDaemonLock(): { pid: number } | undefined {
  try {
    const parsed = JSON.parse(readFileSync(DAEMON_LOCK_FILE, 'utf8')) as { pid?: number }
    if (typeof parsed.pid === 'number' && Number.isInteger(parsed.pid) && parsed.pid > 0) {
      return { pid: parsed.pid }
    }
  } catch {
    // missing or corrupt lock — treat as absent
  }
  return undefined
}

function writeDaemonLock(pid: number): void {
  mkdirSync(DAEMON_DIR, { recursive: true })
  writeFileSync(DAEMON_LOCK_FILE, JSON.stringify({ pid, startedAt: new Date().toISOString() }))
}

// ---------------------------------------------------------------------------
// Daemon lifecycle
// ---------------------------------------------------------------------------

export interface EnsureDaemonResult {
  /** The daemon was already running before this call. */
  alreadyRunning: boolean
  /** Daemon health as reported after it became reachable. */
  health: DaemonHealth
}

/** Test seams for ensureDaemon. */
export interface EnsureDaemonDeps {
  spawnDaemon?: () => number
  startTimeoutMs?: number
}

async function waitForDaemonHealth(timeoutMs: number): Promise<DaemonHealth | undefined> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    const health = await fetchDaemonHealth()
    if (health) return health
    await new Promise((r) => setTimeout(r, DAEMON_POLL_INTERVAL_MS))
  }
  return undefined
}

/**
 * Spawn the daemon detached: own process group (survives harness shutdown and
 * process-group kills), no inherited stdio (log goes to a file), unref'd so
 * this process can exit independently.
 *
 * Daemon environment: the harness session's env, overlaid with
 * ~/.office-live-mcp/daemon.env (KEY=VALUE lines). That file is how users pin
 * daemon settings the harness doesn't know about, e.g. BRIDGE_TLS=1 for
 * HTTPS-sideloaded add-ins.
 */
function readDaemonEnvFile(): Record<string, string> {
  const env: Record<string, string> = {}
  try {
    const lines = readFileSync(join(DAEMON_DIR, 'daemon.env'), 'utf8').split('\n')
    for (const line of lines) {
      const trimmed = line.trim()
      if (!trimmed || trimmed.startsWith('#')) continue
      const eq = trimmed.indexOf('=')
      if (eq <= 0) continue
      env[trimmed.slice(0, eq).trim()] = trimmed.slice(eq + 1).trim()
    }
  } catch {
    // no daemon.env — inherit the plain session env
  }
  return env
}

function spawnDaemon(): number {
  mkdirSync(DAEMON_DIR, { recursive: true })
  const logFd = openSync(DAEMON_LOG_FILE, 'a')
  const entry = process.argv[1]
  if (!entry) throw new Error('Cannot determine server entry point (process.argv[1] is empty)')
  // Source runs need the type-stripping flag; the bundled dist files do not.
  const isTs = entry.endsWith('.ts')
  const args = [...(isTs ? ['--experimental-strip-types'] : []), entry, '--http', '--bridge']
  const child = spawn(process.execPath, args, {
    detached: true,
    stdio: ['ignore', logFd, logFd],
    env: { ...process.env, ...readDaemonEnvFile() },
    // Windows: the detached child would otherwise get its own visible console
    // window popping up on first harness connect.
    windowsHide: true,
  })
  child.unref()
  // Close our copies of the log fd; the daemon inherits its own.
  closeSync(logFd)
  return child.pid ?? -1
}

/**
 * Make sure exactly one daemon is running, starting it if needed.
 * Safe under races: if another wrapper spawns concurrently, the loser's child
 * dies on EADDRINUSE and this wrapper still finds the winner via /health.
 */
export async function ensureDaemon(localVersion: string, deps: EnsureDaemonDeps = {}): Promise<EnsureDaemonResult> {
  const spawnFn = deps.spawnDaemon ?? spawnDaemon
  const startTimeoutMs = deps.startTimeoutMs ?? DAEMON_START_TIMEOUT_MS

  let health = await fetchDaemonHealth()
  if (health) {
    warnOnVersionSkew(health, localVersion, false)
    return { alreadyRunning: true, health }
  }

  const lock = readDaemonLock()
  const lockHeldByLiveProcess = lock !== undefined && isPidAlive(lock.pid)
  if (!lockHeldByLiveProcess) {
    const pid = spawnFn()
    if (pid > 0) writeDaemonLock(pid)
  }

  health = await waitForDaemonHealth(startTimeoutMs)
  if (!health) {
    throw new Error(
      `office-live-mcp daemon did not become healthy within ${startTimeoutMs / 1000}s ` + `(check ${DAEMON_LOG_FILE})`,
    )
  }
  warnOnVersionSkew(health, localVersion, true)
  return { alreadyRunning: false, health }
}

function warnOnVersionSkew(health: DaemonHealth, localVersion: string, started: boolean): void {
  if (health.version && health.version !== localVersion) {
    const verb = started ? 'started' : 'found'
    console.error(
      `[stdio-proxy] Warning: daemon ${verb} is v${health.version} but this wrapper is v${localVersion}. ` +
        'Restart the daemon (office-live-mcp --stop-daemon, then reconnect) to pick up the new version.',
    )
  }
}

/** Stop the daemon regardless of how it was started. Returns true if it was running. */
export async function stopDaemon(): Promise<boolean> {
  try {
    const res = await fetch(`${mcpHttpBase()}/shutdown`, { method: 'POST', signal: AbortSignal.timeout(2_000) })
    return res.ok
  } catch {
    return false
  }
}

// ---------------------------------------------------------------------------
// SSE parsing (the daemon answers POST /mcp with JSON or an SSE stream)
// ---------------------------------------------------------------------------

/** Extract all JSON-RPC messages from an SSE response body. */
export function parseSseMessages(body: string): unknown[] {
  const messages: unknown[] = []
  for (const line of body.split('\n')) {
    const trimmed = line.trim()
    if (!trimmed.startsWith('data:')) continue
    const payload = trimmed.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    try {
      messages.push(JSON.parse(payload))
    } catch {
      // Ignore malformed partials; the client would not be able to use them.
    }
  }
  return messages
}

// ---------------------------------------------------------------------------
// STDIO ↔ HTTP proxy
// ---------------------------------------------------------------------------

interface ProxyOptions {
  stdin?: NodeJS.ReadableStream
  stdout?: NodeJS.WritableStream
  /** Version of this wrapper, used for daemon version-skew warnings. */
  localVersion: string
  /** Injected for tests. */
  ensure?: () => Promise<EnsureDaemonResult>
}

/**
 * Run the stdio proxy: newline-delimited JSON-RPC on stdin → POST /mcp on the
 * daemon → responses (JSON or SSE) back as newline-delimited JSON on stdout.
 * Resolves when stdin ends and all in-flight requests have been answered.
 */
export async function runStdioProxy(options: ProxyOptions): Promise<void> {
  const stdin = options.stdin ?? process.stdin
  const stdout = options.stdout ?? process.stdout
  const ensure = options.ensure ?? (() => ensureDaemon(options.localVersion))
  const { alreadyRunning } = await ensure()
  if (!alreadyRunning) {
    console.error('[stdio-proxy] Daemon started (detached); it stays up after this session ends.')
  }

  let sessionId: string | undefined
  // In-order queue: stdin often delivers several JSON-RPC lines in one chunk,
  // and follow-up requests must not overtake the initialize round-trip that
  // carries the session id. Mirrors the sequential behavior of a plain stdio
  // server.
  let queue: Promise<void> = Promise.resolve()

  function enqueueMessage(raw: string): void {
    queue = queue
      .then(() => handleMessage(raw))
      .catch((err: unknown) => {
        // Transport-level failure: report as a JSON-RPC internal error so the
        // client sees something instead of hanging forever.
        stdout.write(
          `${JSON.stringify({
            jsonrpc: '2.0',
            error: { code: -32603, message: `office-live-mcp proxy error: ${String(err)}` },
            id: null,
          })}\n`,
        )
      })
  }

  async function handleMessage(raw: string): Promise<void> {
    try {
      JSON.parse(raw)
    } catch {
      return // Not valid JSON-RPC; the daemon would reject it too. Drop.
    }
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
    }
    if (sessionId) headers['Mcp-Session-Id'] = sessionId

    const res = await fetch(`${mcpHttpBase()}/mcp`, {
      method: 'POST',
      headers,
      body: raw,
    })

    const newSessionId = res.headers.get('mcp-session-id')
    if (newSessionId) sessionId = newSessionId

    const contentType = res.headers.get('content-type') ?? ''
    const body = await res.text()
    let messages: unknown[]
    if (contentType.includes('text/event-stream')) {
      messages = parseSseMessages(body)
    } else if (body.trim()) {
      messages = [JSON.parse(body)]
    } else {
      messages = [] // e.g. 202 Accepted with no body
    }
    for (const msg of messages) {
      stdout.write(`${JSON.stringify(msg)}\n`)
    }
  }

  return new Promise<void>((resolve, reject) => {
    let buffer = ''
    stdin.setEncoding('utf8')
    stdin.on('data', (chunk: string) => {
      buffer += chunk
      let newlineIndex = buffer.indexOf('\n')
      while (newlineIndex !== -1) {
        const line = buffer.slice(0, newlineIndex).trim()
        buffer = buffer.slice(newlineIndex + 1)
        newlineIndex = buffer.indexOf('\n')
        if (!line) continue
        enqueueMessage(line)
      }
    })
    stdin.on('end', () => {
      queue.then(() => {
        // Best-effort session cleanup so the daemon drops this session.
        if (sessionId) {
          fetch(`${mcpHttpBase()}/mcp`, {
            method: 'DELETE',
            headers: { 'Mcp-Session-Id': sessionId },
          }).catch(() => {})
        }
        resolve()
      }, reject)
    })
    stdin.on('error', reject)
  })
}
