/* OpenCode chat client — talks to the OpenCode HTTP server from inside the taskpane.
   Reverse mode of the bridge: instead of the MCP server driving Office, the add-in
   drives OpenCode (chat) and applies replies to the document via Office.js.
   Plain JS, no build step — loaded with a script tag like app.js. */

var OpenCodeChat = (() => {
  var CONFIG_KEY = 'opencode-chat-config'
  var DEFAULT_URL = 'http://127.0.0.1:4096'

  /* Config — server URL and basic-auth password. The OpenCode background service
     picks a random port and password on every restart, so both are user-supplied. */
  function loadConfig() {
    let raw
    let parsed
    try {
      raw = localStorage.getItem(CONFIG_KEY)
      if (raw) {
        parsed = JSON.parse(raw)
        return { url: parsed.url || DEFAULT_URL, password: parsed.password || '' }
      }
    } catch (_e) {
      // Corrupt config — fall through to defaults
    }
    return { url: DEFAULT_URL, password: '' }
  }

  function saveConfig(config) {
    localStorage.setItem(CONFIG_KEY, JSON.stringify({ url: config.url, password: config.password }))
  }

  /* Authenticated request. OpenCode's local service uses HTTP basic auth
     with a fixed username and a per-service password. */
  function request(config, path, options) {
    var headers = { 'Content-Type': 'application/json' }
    if (config.password) {
      headers.Authorization = `Basic ${btoa(`opencode:${config.password}`)}`
    }
    return fetch(config.url.replace(/\/$/, '') + path, Object.assign({ headers: headers }, options || {})).then(
      (res) => {
        if (!res.ok) {
          return res.text().then((body) => {
            throw new Error(`OpenCode ${res.status}: ${body.slice(0, 300)}`)
          })
        }
        return res.json()
      },
    )
  }

  function createSession(config, title) {
    return request(config, '/api/session', {
      method: 'POST',
      body: JSON.stringify({ title: title || 'Office add-in chat' }),
    }).then((res) => res.data)
  }

  function listSessions(config) {
    return request(config, '/api/session').then((res) => res.data || [])
  }

  function getMessages(config, sessionID) {
    return request(config, `/api/session/${sessionID}/message`).then((res) => res.data || [])
  }

  function sendPrompt(config, sessionID, text) {
    return request(config, `/api/session/${sessionID}/prompt`, {
      method: 'POST',
      body: JSON.stringify({ text: text }),
    }).then((res) => res.data)
  }

  /* Extract displayable text from a message row returned by GET .../message.
     Assistant rows carry a `content` array of typed parts; user rows carry `text`. */
  function messageText(message) {
    if (message.type === 'user') return message.text || ''
    if (!message.content) return ''
    return message.content
      .filter((part) => part.type === 'text')
      .map((part) => part.text)
      .join('\n')
  }

  function isTerminalRow(message) {
    return message.type === 'idle' || message.type === 'step-start'
  }

  /* SSE subscription. EventSource cannot send basic-auth headers, so the stream
     is consumed through fetch + a ReadableStream reader instead. Repeated `types`
     query params filter server-side; `server.connected` always arrives as a marker.
     Live-only: no replay, no auto-reconnect — the caller resubscribes on error. */
  function subscribeEvents(config, types, onEvent) {
    var query = types.map((t) => `types=${encodeURIComponent(t)}`).join('&')
    var headers = {}
    if (config.password) {
      headers.Authorization = `Basic ${btoa(`opencode:${config.password}`)}`
    }
    var controller = new AbortController()
    var stopped = false

    function parseChunk(buffer) {
      var events = []
      var blocks = buffer.split('\n\n')
      var remainder = blocks.pop()
      var i, line
      blocks.forEach((block) => {
        var lines = block.split('\n')
        for (i = 0; i < lines.length; i++) {
          line = lines[i]
          if (line.indexOf('data: ') !== 0) continue
          try {
            events.push(JSON.parse(line.slice(6)))
          } catch (_e) {
            // Heartbeats and partial lines are not JSON — ignore
          }
        }
      })
      return { events: events, remainder: remainder }
    }

    function consume(reader) {
      var buffer = ''
      function pump() {
        return reader.read().then((chunk) => {
          if (stopped) {
            reader.cancel()
            return
          }
          if (chunk.done) {
            if (!stopped) onEvent({ type: 'stream.closed' })
            return
          }
          buffer += new TextDecoder().decode(chunk.value, { stream: true })
          var parsed = parseChunk(buffer)
          buffer = parsed.remainder
          parsed.events.forEach(onEvent)
          return pump()
        })
      }
      return pump()
    }

    fetch(`${config.url.replace(/\/$/, '')}/api/event?${query}`, {
      headers: headers,
      signal: controller.signal,
    })
      .then((res) => {
        if (!res.ok || !res.body) throw new Error(`event stream ${res.status}`)
        return consume(res.body.getReader())
      })
      .catch((err) => {
        if (!stopped) onEvent({ type: 'stream.error', error: String(err?.message ? err.message : err) })
      })

    return {
      stop: () => {
        stopped = true
        controller.abort()
      },
    }
  }

  return {
    loadConfig: loadConfig,
    saveConfig: saveConfig,
    request: request,
    createSession: createSession,
    listSessions: listSessions,
    getMessages: getMessages,
    sendPrompt: sendPrompt,
    messageText: messageText,
    isTerminalRow: isTerminalRow,
    subscribeEvents: subscribeEvents,
    DEFAULT_URL: DEFAULT_URL,
  }
})()

// Expose for the taskpane: chat.js reads the global. The var already lands on
// the global scope, but an explicit window assignment survives bundlers and
// strict module scopes too.
if (typeof window !== 'undefined') {
  window.OpenCodeChat = OpenCodeChat
}
