/* OpenCode chat client — talks to the OpenCode HTTP server from inside the taskpane.
   Reverse mode of the bridge: instead of the MCP server driving Office, the add-in
   drives OpenCode (chat) and applies replies to the document via Office.js.
   Plain JS, no build step — loaded with a script tag like app.js. */

var OpenCodeChat = (() => {
  var CONFIG_KEY = 'opencode-chat-config'
  // Default: the OpenCode proxy built into the bridge server — same origin as
  // the taskpane, so no CORS, no certificate and no password are needed.
  var DEFAULT_URL = '/opencode'

  /* Config — server URL and basic-auth password. The OpenCode background service
     picks a random port and password on every restart, so both are user-supplied. */
  function loadConfig() {
    let raw
    let parsed
    try {
      raw = localStorage.getItem(CONFIG_KEY)
      if (raw) {
        parsed = JSON.parse(raw)
        const url = parsed.url || DEFAULT_URL
        // Migrate configs from the retired standalone proxy (port 4096) to the
        // bridge's built-in same-origin route.
        const migrated = /^https?:\/\/127\.0\.0\.1:4096\/?$/i.test(url) ? DEFAULT_URL : url
        return { url: migrated, password: parsed.password || '' }
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
  /* Resolve the request base: an absolute URL (direct connection) or a
     same-origin path like the default '/opencode' (bridge proxy). */
  function base(config) {
    var url = (config.url || DEFAULT_URL).trim()
    if (!url) url = DEFAULT_URL
    return url.replace(/\/$/, '')
  }

  function request(config, path, options) {
    var headers = { 'Content-Type': 'application/json' }
    if (config.password) {
      headers.Authorization = `Basic ${btoa(`opencode:${config.password}`)}`
    }
    return fetch(base(config) + path, Object.assign({ headers: headers }, options || {})).then((res) => {
      if (!res.ok) {
        return res.text().then((body) => {
          throw new Error(`OpenCode ${res.status}: ${body.slice(0, 300)}`)
        })
      }
      return res.json()
    })
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

  /* Map pending attachments to the prompt API's `files` array. Images travel
     inline as data URIs — the taskpane has no server-side storage to link to. */
  function promptFiles(files) {
    return (files || [])
      .filter((f) => f?.dataUrl)
      .map((f) => ({ uri: f.dataUrl, source: { type: 'inline' }, name: f.name }))
  }

  function sendPrompt(config, sessionID, text, files) {
    var body = { text: text }
    var mapped = promptFiles(files)
    if (mapped.length) body.files = mapped
    return request(config, `/api/session/${sessionID}/prompt`, {
      method: 'POST',
      body: JSON.stringify(body),
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

  /* Images the model rendered (file parts with an image mime) as data URIs,
     plus images a tool wrote to disk (paths found in tool output text) served
     through the bridge's /file endpoint on the taskpane's own origin. */
  function messageImages(message) {
    if (!message.content) return []
    var images = (message.content || [])
      .filter(
        (part) =>
          part.type === 'file' && typeof part.mime === 'string' && part.mime.indexOf('image/') === 0 && part.data,
      )
      .map((part) => ({ mime: part.mime, name: part.name || 'image', src: `data:${part.mime};base64,${part.data}` }))
    var origin = typeof window !== 'undefined' ? window.location.origin : ''
    ;(message.content || []).forEach((part) => {
      if (part.type !== 'tool') return
      var state = part.state || {}
      ;(state.content || []).forEach((block) => {
        if (block.type !== 'text' || !block.text) return
        // Mirrors server/file-serve.ts imagePathsInText (kept in sync).
        var matches = block.text.match(/(?:~?\/)[\w.@/-]+\.(?:png|jpe?g|gif|webp|svg)/gi) || []
        matches.forEach((match) => {
          if (match.startsWith('/') && block.text[block.text.indexOf(match) - 1] === '.') return
          if (images.some((existing) => existing.src.indexOf(encodeURIComponent(match)) !== -1)) return
          images.push({
            mime: 'image/*',
            name: match.split('/').pop() || match,
            src: `${origin}/file?path=${encodeURIComponent(match)}`,
          })
        })
      })
    })
    return images
  }

  function isTerminalRow(message) {
    return message.type === 'idle' || message.type === 'step-start'
  }

  /* Map raw network errors to something a user can act on. WKWebView reports
     unreachable servers as "Load failed", Chrome as "Failed to fetch". */
  function friendlyError(err) {
    const raw = String(err?.message ? err.message : err)
    if (/failed to fetch|load failed|networkerror/i.test(raw)) {
      return 'OpenCode not reachable — is the bridge running with the OpenCode proxy? (URL "/opencode" in Connection settings)'
    }
    return raw
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

    fetch(`${base(config)}/api/event?${query}`, {
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
    promptFiles: promptFiles,
    messageText: messageText,
    messageImages: messageImages,
    isTerminalRow: isTerminalRow,
    friendlyError: friendlyError,
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
