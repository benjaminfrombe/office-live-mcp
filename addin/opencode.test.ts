import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { runInNewContext } from 'node:vm'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

interface Harness {
  context: Record<string, unknown>
  fetchCalls: { url: string; init: RequestInit | undefined }[]
  respond: (status: number, body: unknown) => void
}

function harness(): Harness {
  const context: Record<string, unknown> = {}
  const fetchCalls: { url: string; init: RequestInit | undefined }[] = []
  let responder: ((status: number, body: unknown) => void) | null = null

  const storage: Storage = {
    store: new Map<string, string>(),
    getItem(key: string) {
      return (this as unknown as { store: Map<string, string> }).store.get(key) ?? null
    },
    setItem(key: string, value: string) {
      ;(this as unknown as { store: Map<string, string> }).store.set(key, value)
    },
  } as unknown as Storage
  context.localStorage = storage
  context.btoa = (input: string) => Buffer.from(input, 'binary').toString('base64')
  context.TextDecoder = TextDecoder
  context.AbortController = AbortController
  context.fetch = (url: string, init?: RequestInit) => {
    fetchCalls.push({ url, init })
    return new Promise((resolve, _reject) => {
      responder = (status: number, body: unknown) => {
        resolve(
          new Response(typeof body === 'string' ? body : JSON.stringify(body), {
            status,
            headers: { 'Content-Type': 'application/json' },
          }),
        )
      }
    })
  }

  const source = readFileSync(join(__dirname, 'opencode.js'), 'utf8')
  runInNewContext(source, context)

  return {
    context,
    fetchCalls,
    respond: (status, body) => {
      if (!responder) throw new Error('no pending fetch')
      responder(status, body)
    },
  }
}

beforeEach(() => {
  vi.useFakeTimers()
})
afterEach(() => {
  vi.useRealTimers()
})

describe('opencode.js config', () => {
  it('defaults to the bridge proxy route with empty password', () => {
    const h = harness()
    const config = (h.context.OpenCodeChat as { loadConfig(): { url: string; password: string } }).loadConfig()
    expect(config.url).toBe('/opencode')
    expect(config.password).toBe('')
  })

  it('round-trips a saved config through localStorage', () => {
    const h = harness()
    const client = h.context.OpenCodeChat as {
      saveConfig: (c: { url: string; password: string }) => void
      loadConfig(): { url: string; password: string }
    }
    client.saveConfig({ url: 'http://127.0.0.1:49374', password: 'secret' })
    expect(client.loadConfig()).toEqual({ url: 'http://127.0.0.1:49374', password: 'secret' })
  })

  it('falls back to defaults on corrupt stored config', () => {
    const h = harness()
    const store = (h.context.localStorage as { store: Map<string, string> }).store
    store.set('opencode-chat-config', '{not json')
    const client = h.context.OpenCodeChat as { loadConfig(): { url: string; password: string } }
    expect(client.loadConfig().url).toBe('/opencode')
  })
})

describe('opencode.js requests', () => {
  it('sends basic auth and unwraps the data envelope on createSession', async () => {
    const h = harness()
    const client = h.context.OpenCodeChat as {
      saveConfig: (c: { url: string; password: string }) => void
      createSession: (c: unknown, title: string) => Promise<{ id: string }>
    }
    client.saveConfig({ url: 'http://127.0.0.1:49374/', password: 'pw' })
    const pending = client.createSession({ url: 'http://127.0.0.1:49374/', password: 'pw' }, 'test')
    h.respond(200, { data: { id: 'ses_1' } })
    await expect(pending).resolves.toEqual({ id: 'ses_1' })
    expect(h.fetchCalls[0].url).toBe('http://127.0.0.1:49374/api/session')
    const headers = h.fetchCalls[0].init?.headers as Record<string, string>
    expect(headers.Authorization).toBe(`Basic ${Buffer.from('opencode:pw').toString('base64')}`)
    expect(JSON.parse(String(h.fetchCalls[0].init?.body))).toEqual({ title: 'test' })
  })

  it('throws a readable error on non-2xx responses', async () => {
    const h = harness()
    const client = h.context.OpenCodeChat as { request: (c: unknown, p: string) => Promise<unknown> }
    const pending = client.request({ url: 'http://x', password: '' }, '/api/info')
    h.respond(401, { message: 'Authentication required' })
    await expect(pending).rejects.toThrow('OpenCode 401')
  })

  it('sends the prompt text on sendPrompt', async () => {
    const h = harness()
    const client = h.context.OpenCodeChat as {
      sendPrompt: (c: unknown, sid: string, text: string) => Promise<unknown>
    }
    const pending = client.sendPrompt({ url: 'http://x', password: '' }, 'ses_9', 'hello')
    h.respond(200, { data: { id: 'msg_1' } })
    await pending
    expect(h.fetchCalls[0].url).toBe('http://x/api/session/ses_9/prompt')
    expect(JSON.parse(String(h.fetchCalls[0].init?.body))).toEqual({ text: 'hello' })
  })
})

describe('opencode.js message helpers', () => {
  const client = () =>
    harness().context.OpenCodeChat as {
      messageText: (m: Record<string, unknown>) => string
      isTerminalRow: (m: Record<string, unknown>) => boolean
      friendlyError: (e: unknown) => string
    }

  it('extracts user text', () => {
    expect(client().messageText({ type: 'user', text: 'hi' })).toBe('hi')
  })

  it('joins assistant text parts and skips non-text parts', () => {
    expect(
      client().messageText({
        type: 'assistant',
        content: [{ type: 'step-start' }, { type: 'text', text: 'line one' }, { type: 'text', text: 'line two' }],
      }),
    ).toBe('line one\nline two')
  })

  it('flags terminal rows', () => {
    expect(client().isTerminalRow({ type: 'idle' })).toBe(true)
    expect(client().isTerminalRow({ type: 'assistant' })).toBe(false)
  })

  it('maps network failures to an actionable message', () => {
    expect(client().friendlyError(new Error('Load failed'))).toMatch(/not reachable/)
    expect(client().friendlyError(new Error('Failed to fetch'))).toMatch(/not reachable/)
    expect(client().friendlyError(new Error('OpenCode 401: nope'))).toBe('OpenCode 401: nope')
  })
})

describe('opencode.js SSE subscription', () => {
  async function sseHarness() {
    vi.useRealTimers()
    const h = harness()
    const events: Record<string, unknown>[] = []
    let push: ((chunk: string) => void) | null = null
    let close: (() => void) | null = null
    const client = h.context.OpenCodeChat as {
      subscribeEvents: (c: unknown, types: string[], onEvent: (e: Record<string, unknown>) => void) => { stop(): void }
    }
    const encoder = new TextEncoder()
    const stream = new ReadableStream({
      start(controller) {
        push = (chunk: string) => controller.enqueue(encoder.encode(chunk))
        close = () => controller.close()
      },
    })
    // Fetch in harness resolves via responder — override with a stream response
    ;(h.context as { fetch: (url: string, init?: RequestInit) => Promise<Response> }).fetch = () => {
      return Promise.resolve(new Response(stream, { status: 200 }))
    }
    const sub = client.subscribeEvents({ url: 'http://x', password: '' }, ['session.text.delta'], (e) => events.push(e))
    await new Promise((r) => setTimeout(r, 5))
    return { h, events, push: push as unknown as (c: string) => void, close: close as unknown as () => void, sub }
  }

  it('parses data lines into events across chunk boundaries', async () => {
    const sse = await sseHarness()
    sse.push('data: {"type":"session.text.delta","data":{"delta":"HE')
    sse.push('LLO"}}\n\n')
    await vi.waitFor(() => expect(sse.events).toEqual([{ type: 'session.text.delta', data: { delta: 'HELLO' } }]))
  })

  it('ignores heartbeat lines and malformed data', async () => {
    const sse = await sseHarness()
    sse.push(': heartbeat\n\ndata: not-json\n\ndata: {"type":"session.step.ended"}\n\n')
    await vi.waitFor(() => expect(sse.events).toEqual([{ type: 'session.step.ended' }]))
  })

  it('stop() aborts the stream and drops further events', async () => {
    const sse = await sseHarness()
    sse.sub.stop()
    sse.push('data: {"type":"session.text.delta"}\n\n')
    await new Promise((r) => setTimeout(r, 10))
    expect(sse.events).toEqual([])
  })
})
