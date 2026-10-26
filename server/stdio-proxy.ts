// STDIO → HTTP daemon proxy.
//
// When a harness (Claude Code, OpenCode, Claude Desktop, ...) loads this MCP
// server over stdio, this module makes sure a single shared daemon is running
// (`--http --bridge`, detached) and then proxies newline-delimited JSON-RPC
// from stdin/stdout to the daemon's HTTP MCP endpoint.
//
// Why a proxy instead of running the bridge in-process: the bridge is a
// single-instance server (one process binds :8790/:8791/:8792) that must stay
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
const DAEMON_SHUTDOWN_WAIT_MS = 3_000

/** Base URL of the daemon's HTTP MCP server (read per call so tests can point it elsewhere). */
function mcpHttpBase(): string {
  return `http://127.0.0.1:${Number(process.env.MCP_PORT) || 8792}`
}

/**
 * Base URL of the bridge server. The bridge serves the Office add-in and —
 * in every build since MCP-on-the-bridge — also exposes /mcp, so any process
 * that owns the bridge can serve MCP to other harness sessions.
 */
function bridgeHttpBase(): string {
  return `http://127.0.0.1:${Number(process.env.BRIDGE_PORT) || 8790}`
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
  /** True on bridge /health responses of builds that serve /mcp on the bridge port. */
  mcp?: boolean
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

/**
 * Probe the bridge's /health endpoint. Only matches when the listener is an
 * office-live-mcp bridge that also serves /mcp on the same port (health.mcp
 * === true), so a foreign app that happens to sit on the port is ignored.
 * Plain-HTTP only: TLS setups always run a daemon, whose MCP endpoint is
 * reachable over plain HTTP on the MCP port and found first.
 */
export async function fetchBridgeHealth(timeoutMs = HEALTH_TIMEOUT_MS): Promise<DaemonHealth | undefined> {
  try {
    const res = await fetch(`${bridgeHttpBase()}/health`, { signal: AbortSignal.timeout(timeoutMs) })
    if (!res.ok) return undefined
    const health = (await res.json()) as DaemonHealth
    return health.mcp === true ? health : undefined
  } catch {
    return undefined
  }
}

/**
 * Find any live office-live-mcp MCP endpoint: the shared daemon's HTTP server
 * (MCP port) or — when no daemon exists — a bridge that serves /mcp itself
 * (e.g. Claude Desktop's in-process extension bridge). Returns the base URL
 * to POST JSON-RPC to, or undefined when neither answers.
 */
export async function discoverMcpEndpoint(): Promise<string | undefined> {
  const daemon = await fetchDaemonHealth()
  if (daemon) return mcpHttpBase()
  const bridge = await fetchBridgeHealth()
  if (bridge) return bridgeHttpBase()
  return undefined
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
  /** Base URL of the MCP endpoint to proxy to (daemon MCP port, or a bridge that serves /mcp). */
  base: string
}

/** Test seams for ensureDaemon. */
export interface EnsureDaemonDeps {
  spawnDaemon?: () => number
  startTimeoutMs?: number
  /** Overrides stopDaemon() for tests. */
  stopDaemonFn?: () => Promise<boolean>
  /** How long to wait for the stale daemon's endpoint to disappear after a shutdown. */
  shutdownWaitMs?: number
}

async function waitForDaemonHealth(timeoutMs: number): Promise<{ health: DaemonHealth; base: string } | undefined> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    const health = await fetchDaemonHealth()
    if (health) return { health, base: mcpHttpBase() }
    // Another session may have won the race and started an in-process bridge
    // (e.g. inside Claude Desktop's sandbox, where daemon spawns die). Such a
    // bridge serves /mcp itself — converge onto it instead of waiting out the
    // full daemon timeout.
    const bridge = await fetchBridgeHealth()
    if (bridge) return { health: bridge, base: bridgeHttpBase() }
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
  const shutdownWaitMs = deps.shutdownWaitMs ?? DAEMON_SHUTDOWN_WAIT_MS
  const stopFn = deps.stopDaemonFn ?? stopDaemon

  // A bridge that serves /mcp counts as "already running" too: it owns the
  // add-in, so no daemon should be spawned (it would die on EADDRINUSE).
  const existingBase = await discoverMcpEndpoint()
  if (existingBase) {
    const existingHealth = await fetchDaemonHealth()
    if (existingHealth && versionsMatch(existingHealth.version, localVersion)) {
      return { alreadyRunning: true, health: existingHealth, base: existingBase }
    }
    // Version skew: the running daemon predates this build, and the daemon
    // holds all tool logic — a rebuilt dist on disk changes nothing until the
    // process restarts. Restart it instead of warning: shutdown, wait for the
    // endpoint to disappear, then fall through to a fresh spawn below.
    console.error(
      `[stdio-proxy] Daemon is v${existingHealth?.version ?? 'unknown'} but this wrapper is v${localVersion}. ` +
        'Restarting the daemon to pick up the new version.',
    )
    await stopFn()
    await waitForEndpointToDisappear(shutdownWaitMs)
  }

  const lock = readDaemonLock()
  const lockHeldByLiveProcess = lock !== undefined && isPidAlive(lock.pid)
  if (!lockHeldByLiveProcess) {
    const pid = spawnFn()
    if (pid > 0) writeDaemonLock(pid)
  }

  const found = await waitForDaemonHealth(startTimeoutMs)
  if (!found) {
    throw new Error(
      `office-live-mcp daemon did not become healthy within ${startTimeoutMs / 1000}s ` + `(check ${DAEMON_LOG_FILE})`,
    )
  }
  warnOnVersionSkew(found.health, localVersion, true)
  return { alreadyRunning: false, health: found.health, base: found.base }
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

/**
 * A daemon without a version field is treated as matching: very old daemons
 * and foreign bridges do not report one, and restarting those from the proxy
 * would do more harm than good.
 */
function versionsMatch(daemonVersion: string | undefined, localVersion: string): boolean {
  return !daemonVersion || daemonVersion === localVersion
}

/** Poll until the MCP endpoint stops answering (best-effort, bounded). */
async function waitForEndpointToDisappear(timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (!(await discoverMcpEndpoint())) return
    await new Promise((r) => setTimeout(r, DAEMON_POLL_INTERVAL_MS))
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
  const { alreadyRunning, base } = await ensure()
  if (!alreadyRunning) {
    console.error('[stdio-proxy] Daemon started (detached); it stays up after this session ends.')
  }

  let sessionId: string | undefined
  // Until the daemon has issued a session id, requests run strictly in order:
  // stdin often delivers several JSON-RPC lines in one chunk, and follow-ups
  // must not overtake the initialize round-trip that carries the session id.
  // Once the session exists, requests run concurrently so one slow tool call
  // cannot hold up every later call (head-of-line blocking) or a cancellation.
  let queue: Promise<void> = Promise.resolve()
  const inflight = new Set<Promise<void>>()

  function sendMessage(raw: string): Promise<void> {
    return handleMessage(raw).catch((err: unknown) => {
      // Transport-level failure: answer the failed request under its own id so
      // the client sees an error instead of waiting for its own timeout.
      // Notifications carry no id and get no response.
      let id: unknown
      try {
        id = (JSON.parse(raw) as { id?: unknown }).id
      } catch {
        // unparsable input was already dropped in handleMessage
      }
      if (id === undefined) {
        console.error(`[stdio-proxy] notification failed: ${String(err)}`)
        return
      }
      stdout.write(
        `${JSON.stringify({
          jsonrpc: '2.0',
          error: { code: -32603, message: `office-live-mcp proxy error: ${String(err)}` },
          id,
        })}\n`,
      )
    })
  }

  function enqueueMessage(raw: string): void {
    let task: Promise<void>
    if (sessionId === undefined) {
      queue = queue.then(() => sendMessage(raw))
      task = queue
    } else {
      task = sendMessage(raw)
    }
    inflight.add(task)
    void task.finally(() => inflight.delete(task))
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

    const res = await fetch(`${base}/mcp`, {
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
      Promise.all([queue, ...inflight]).then(() => {
        // Best-effort session cleanup so the daemon drops this session.
        if (sessionId) {
          fetch(`${base}/mcp`, {
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
