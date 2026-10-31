/* OpenCode chat UI — taskpane mode that chats with the OpenCode server and
   applies replies to the live document via Office.js. Runs next to the bridge
   mode in the same taskpane; the bridge keeps working while the chat is open. */

var chatState = {
  config: null,
  sessionID: null,
  events: null, // active SSE subscription
  streamingText: '', // partial assistant text while streaming
  busy: false,
  pendingSelections: [], // document selections kept as context for the next message
  pendingAttachments: [], // pasted images kept as data URIs for the next message
  reasoningText: '', // live assistant reasoning buffer for the strip
  reasoningActive: false, // true while the current step is still streaming reasoning
  renderedIds: [], // message ids already on screen (incremental rendering)
}

var EVENT_TYPES = [
  'session.step.started',
  'session.text.delta',
  'session.text.ended',
  'session.reasoning.started',
  'session.reasoning.delta',
  'session.reasoning.ended',
  'session.step.ended',
  'session.step.failed',
  'session.execution.started',
  'session.execution.succeeded',
  'session.execution.failed',
  'session.execution.interrupted',
]

function initChat() {
  chatState.config = OpenCodeChat.loadConfig()
  document.getElementById('oc-server-url').value = chatState.config.url
  document.getElementById('oc-password').value = chatState.config.password
  // Only show the password field when a direct (non-proxy) connection is stored
  if (chatState.config.password) document.getElementById('oc-password').classList.remove('hidden')

  document.getElementById('oc-save-config').addEventListener('click', () => {
    chatState.config = {
      url: document.getElementById('oc-server-url').value.trim() || OpenCodeChat.DEFAULT_URL,
      password: document.getElementById('oc-password').value,
    }
    OpenCodeChat.saveConfig(chatState.config)
    setChatStatus('Config saved — reconnecting...')
    refreshSessions()
    focusInput()
  })

  document.getElementById('oc-toggle-advanced').addEventListener('click', () => {
    const password = document.getElementById('oc-password')
    password.classList.toggle('hidden')
    if (!password.classList.contains('hidden')) {
      password.focus()
    } else {
      focusInput()
    }
  })

  document.getElementById('oc-new-session').addEventListener('click', () => {
    OpenCodeChat.createSession(chatState.config, 'Office add-in chat').then((session) => {
      selectSession(session.id)
      refreshSessions()
      focusInput()
    }, chatError)
  })

  document.getElementById('oc-send').addEventListener('click', sendChatMessage)
  const input = document.getElementById('oc-input')
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      sendChatMessage()
    }
  })
  input.addEventListener('input', autoGrowInput)
  input.addEventListener('paste', onInputPaste)

  document.getElementById('oc-reasoning-strip').addEventListener('click', () => {
    var strip = document.getElementById('oc-reasoning-strip')
    strip.dataset.expanded = strip.dataset.expanded === '1' ? '' : '1'
    updateReasoningStrip()
  })

  document.getElementById('oc-use-selection').addEventListener('click', () => {
    runLocalOffice(
      'const out = { text: "", slide: null, shapes: [] };\n' +
        'const raw = await new Promise((resolve) => ' +
        'Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => resolve(r.value)));\n' +
        'out.text = String(raw || "");\n' +
        'const pres = context.presentation;\n' +
        'try {\n' +
        '  const shapes = pres.getSelectedShapes();\n' +
        '  shapes.load("items");\n' +
        '  await context.sync();\n' +
        '  for (const s of shapes.items) {\n' +
        '    let text = "";\n' +
        '    try {\n' +
        '      s.load("name,id,left,top,width,height,textFrame/hasText");\n' +
        '      await context.sync();\n' +
        '      if (s.textFrame.hasText) {\n' +
        '        s.textFrame.textRange.load("text");\n' +
        '        await context.sync();\n' +
        '        text = s.textFrame.textRange.text;\n' +
        '      }\n' +
        '    } catch (e) { continue; }\n' +
        '    out.shapes.push({ name: s.name, id: s.id, left: s.left, top: s.top, width: s.width, height: s.height, text: String(text || "") });\n' +
        '  }\n' +
        '} catch (e) {}\n' +
        'try {\n' +
        '  const slides = pres.getSelectedSlides();\n' +
        '  slides.load("items");\n' +
        '  await context.sync();\n' +
        '  if (slides.items.length) out.slide = slides.items.map((x) => x.index + 1);\n' +
        '} catch (e) {}\n' +
        'return out;',
    ).then((out) => {
      // Prefer the exact text-range text; fall back to per-shape text so
      // shape-only selections (no text) still carry their metadata.
      let text = String(out && out.text ? out.text : '')
        .replace(/\r/g, '')
        .replace(/\n{3,}/g, '\n\n')
        .trim()
      const shapes = (out && out.shapes) || []
      if (!text && shapes.length) {
        text = shapes
          .map((s) => s.text.trim())
          .filter(Boolean)
          .join('\n\n')
      }
      if (!text && shapes.length) text = '(shape without text)'
      if (text && shapes.length) {
        // Office can return far more than the visible selection (whole shapes,
        // surrounding paragraphs). Cap it so prompts stay focused.
        const MAX_SELECTION_CHARS = 1200
        if (text.length > MAX_SELECTION_CHARS) {
          text = `${text.slice(0, MAX_SELECTION_CHARS)}\n[selection truncated, ${text.length} chars total]`
        }
        const entry = { text, slide: out.slide, shapes }
        const key = (e) => `${e.shapes.map((s) => s.id).join(',')}|${e.text}`
        const dupe = chatState.pendingSelections.some((s) => key(s) === key(entry))
        if (!dupe) chatState.pendingSelections.push(entry)
        renderSelectionChips()
      }
      focusInput()
    }, chatError)
  })

  document.getElementById('oc-insert-reply').addEventListener('click', () => {
    var text = lastAssistantText()
    if (!text) {
      setChatStatus('No assistant reply to insert yet')
      focusInput()
      return
    }
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setChatStatus(`Insert failed: ${result.error ? result.error.message : 'unknown'}`)
      } else {
        setChatStatus('Reply inserted at selection')
      }
      focusInput()
    })
  })

  refreshSessions()
  autoGrowInput()

  // Auto-open the chat tab when OpenCode is reachable; stay on Bridge otherwise
  OpenCodeChat.request(chatState.config, '/api/info', { signal: AbortSignal.timeout(3000) }).then(
    () => {
      showMode('chat')
      focusInput()
      scrollMessages()
    },
    () => {
      /* OpenCode not reachable — keep the Bridge tab */
    },
  )
}

function setChatStatus(text) {
  var el = document.getElementById('oc-status')
  if (el) el.textContent = text
}

function chatError(err) {
  console.error('OpenCode chat error:', err)
  setChatStatus(OpenCodeChat.friendlyError(err))
}

function refreshSessions() {
  OpenCodeChat.listSessions(chatState.config).then((sessions) => {
    var select = document.getElementById('oc-session')
    select.innerHTML = ''
    sessions.slice(0, 25).forEach((session) => {
      var option = document.createElement('option')
      option.value = session.id
      option.textContent = session.title || session.id
      select.appendChild(option)
    })
    if (chatState.sessionID) select.value = chatState.sessionID
    if (select.value) selectSession(select.value)
  }, chatError)
}

function selectSession(sessionID) {
  stopEventStream()
  chatState.sessionID = sessionID
  chatState.streamingText = ''
  chatState.renderedIds = []
  chatState.reasoningText = ''
  chatState.reasoningActive = false
  updateReasoningStrip()
  document.getElementById('oc-session').value = sessionID
  renderFromRows([])
  setChatStatus('Loading messages...')
  OpenCodeChat.getMessages(chatState.config, sessionID).then(renderFromRows, chatError)
  startEventStream()
}

/* Event stream — deltas append to streamingText; terminal events trigger a
   full message refetch so the rendered history always matches the server. */
/* Event stream — deltas append to streamingText; terminal events trigger a
   full message refetch so the rendered history always matches the server. */
var REFETCH_DEBOUNCE_MS = 150
var refetchTimer = null

function scheduleRefetch(sessionID) {
  // step.ended and execution.succeeded often fire together — one refetch is enough
  if (refetchTimer) clearTimeout(refetchTimer)
  refetchTimer = setTimeout(() => {
    refetchTimer = null
    if (chatState.sessionID === sessionID) {
      OpenCodeChat.getMessages(chatState.config, sessionID).then(renderFromRows, chatError)
    }
  }, REFETCH_DEBOUNCE_MS)
}

function startEventStream() {
  var sessionID = chatState.sessionID
  if (!sessionID) return
  chatState.events = OpenCodeChat.subscribeEvents(chatState.config, EVENT_TYPES, (event) => {
    if (event.type === 'stream.closed' || event.type === 'stream.error') {
      if (chatState.sessionID === sessionID) {
        // Exponential backoff — an absent server must not cause a tight retry loop
        chatState.streamRetries = (chatState.streamRetries || 0) + 1
        const delay = Math.min(2000 * 2 ** (chatState.streamRetries - 1), 30000)
        if (event.type === 'stream.error') setChatStatus(OpenCodeChat.friendlyError(event.error))
        setTimeout(startEventStream, delay)
      }
      return
    }
    if (event.type === 'server.connected') {
      chatState.streamRetries = 0
      return
    }
    var data = event.data || {}
    if (data.sessionID !== sessionID) return
    if (event.type === 'session.text.delta') {
      chatState.streamingText += data.delta || ''
      renderStreaming()
    } else if (event.type === 'session.text.ended') {
      chatState.streamingText = ''
      renderStreaming()
    } else if (event.type === 'session.reasoning.delta') {
      chatState.reasoningText += data.delta || ''
      updateReasoningStrip()
    } else if (event.type === 'session.reasoning.ended') {
      chatState.reasoningActive = false
      updateReasoningStrip()
    } else if (event.type === 'session.step.started') {
      // A new step starts a fresh reasoning turn; keep the previous one only
      // until this point so the strip never mixes two steps.
      chatState.reasoningText = ''
      chatState.reasoningActive = true
      updateReasoningStrip()
      setChatStatus('Assistant is working...')
    } else if (
      event.type === 'session.step.ended' ||
      event.type === 'session.step.failed' ||
      event.type === 'session.execution.succeeded' ||
      event.type === 'session.execution.failed' ||
      event.type === 'session.execution.interrupted'
    ) {
      chatState.streamingText = ''
      renderStreaming()
      chatState.reasoningActive = false
      updateReasoningStrip()
      setChatStatus('')
      scheduleRefetch(sessionID)
    }
  })
}

function stopEventStream() {
  if (chatState.events) {
    chatState.events.stop()
    chatState.events = null
  }
}

function sendChatMessage() {
  const input = document.getElementById('oc-input')
  const typed = input.value.trim()
  const attachments = chatState.pendingAttachments
  if (
    (!typed && !chatState.pendingSelections.length && !attachments.length) ||
    !chatState.sessionID ||
    chatState.busy
  ) {
    focusInput()
    return
  }
  const context = chatState.pendingSelections.length ? `${selectionContextBlocks()}\n\n` : ''
  const text = `${context}${typed}`
  chatState.busy = true
  input.value = ''
  chatState.pendingSelections = []
  renderSelectionChips()
  autoGrowInput()
  setChatStatus('Sending...')
  appendMessage(
    'user',
    typed ||
      (attachments.length
        ? `(${attachments.length} image${attachments.length > 1 ? 's' : ''})`
        : '(document selection)'),
    true,
  )
  autoGrowInput()
  focusInput()
  OpenCodeChat.sendPrompt(chatState.config, chatState.sessionID, text, attachments).then(
    () => {
      chatState.busy = false
      // Attachments are cleared only on success so a failed send can be retried.
      chatState.pendingAttachments = []
      renderAttachmentChips()
      setChatStatus('')
    },
    (err) => {
      chatState.busy = false
      chatError(err)
    },
  )
}

/* Keep the input focused: every action returns focus so the user can keep
   typing (or press Enter) without clicking back into the field. */
function focusInput() {
  const input = document.getElementById('oc-input')
  if (input && !document.getElementById('mode-chat').classList.contains('hidden')) {
    input.focus()
  }
}

/* Selection context chips — each captured selection is one removable chip;
   they are prepended to the prompt only when the message is sent. */
function renderSelectionChips() {
  const container = document.getElementById('oc-selection-chips')
  if (!container) return
  container.innerHTML = ''
  chatState.pendingSelections.forEach((pending, index) => {
    const chip = document.createElement('div')
    chip.className = 'selection-chip'
    const label = document.createElement('span')
    const text = pending.text
    const preview = text.slice(0, 40).replace(/\s+/g, ' ')
    const shapeNames = (pending.shapes || []).map((s) => s.name || 'shape')
    const shapeLabel = shapeNames.length
      ? shapeNames[0] + (shapeNames.length > 1 ? ` +${shapeNames.length - 1}` : '')
      : 'shape'
    const where = pending.slide ? ` · slide ${pending.slide} · ${shapeLabel}` : ''
    label.textContent = `Selection${where} · ${text.length} chars · ${preview}${text.length > 40 ? '…' : ''}`
    const clear = document.createElement('button')
    clear.type = 'button'
    clear.title = 'Remove this selection'
    clear.textContent = '✕'
    clear.addEventListener('click', () => {
      chatState.pendingSelections.splice(index, 1)
      renderSelectionChips()
      focusInput()
    })
    chip.appendChild(label)
    chip.appendChild(clear)
    container.appendChild(chip)
  })
}

/* Pasted image attachments — image items from the clipboard become pending
   attachments rendered as thumbnail chips; WKWebView may not deliver image
   paste items at all, in which case this silently does nothing. */
var ATTACHMENT_MAX_BASE64_BYTES = 8 * 1024 * 1024 // base64 payload cap

function onInputPaste(event) {
  var items = event.clipboardData ? event.clipboardData.items : null
  if (!items) return
  var found = false
  for (let i = 0; i < items.length; i++) {
    const type = items[i].type || ''
    if (type.indexOf('image/') !== 0) continue
    const file = items[i].getAsFile ? items[i].getAsFile() : null
    if (!file) continue
    found = true
    readAttachment(file)
  }
  if (found) event.preventDefault()
}

function readAttachment(file) {
  var reader = new FileReader()
  reader.onload = () => {
    var dataUrl = String(reader.result || '')
    if (dataUrl.length > ATTACHMENT_MAX_BASE64_BYTES) {
      setChatStatus(`Image too large — over 8 MB base64 (${file.name || 'pasted image'})`)
      return
    }
    chatState.pendingAttachments.push({ dataUrl: dataUrl, name: file.name || 'pasted image', size: file.size || 0 })
    renderAttachmentChips()
    setChatStatus('')
    focusInput()
  }
  reader.onerror = () => setChatStatus('Could not read the pasted image')
  reader.readAsDataURL(file)
}

function renderAttachmentChips() {
  const container = document.getElementById('oc-attachment-chips')
  if (!container) return
  container.innerHTML = ''
  chatState.pendingAttachments.forEach((attachment, index) => {
    const chip = document.createElement('div')
    chip.className = 'attachment-chip'
    const thumb = document.createElement('img')
    thumb.src = attachment.dataUrl
    thumb.alt = attachment.name
    const label = document.createElement('span')
    label.textContent = `${attachment.name} · ${Math.round(attachment.size / 1024)} kB`
    const clear = document.createElement('button')
    clear.type = 'button'
    clear.title = 'Remove this image'
    clear.textContent = '✕'
    clear.addEventListener('click', () => {
      chatState.pendingAttachments.splice(index, 1)
      renderAttachmentChips()
      focusInput()
    })
    chip.appendChild(thumb)
    chip.appendChild(label)
    chip.appendChild(clear)
    container.appendChild(chip)
  })
}

/* Reasoning strip — a thin ticker above the messages showing the assistant's
   live reasoning; after the step it stays as a one-line collapsed summary. */
function updateReasoningStrip() {
  var strip = document.getElementById('oc-reasoning-strip')
  if (!strip) return
  var text = chatState.reasoningText
  if (!text) {
    strip.hidden = true
    strip.dataset.expanded = ''
    return
  }
  strip.hidden = false
  var label = document.getElementById('oc-reasoning-label')
  var body = document.getElementById('oc-reasoning-text')
  body.textContent = text
  var expanded = false
  if (chatState.reasoningActive) {
    strip.classList.add('live')
    label.textContent = 'Reasoning…'
    body.scrollTop = body.scrollHeight
  } else {
    strip.classList.remove('live')
    expanded = strip.dataset.expanded === '1'
    label.textContent = `Reasoning · ${text.length} chars · click to ${expanded ? 'collapse' : 'expand'}`
    body.classList.toggle('expanded', expanded)
    if (expanded) body.scrollTop = body.scrollHeight
  }
}

/* Build the context blocks for all pending selections. Points are 1/72 inch. */
function selectionContextBlocks() {
  return chatState.pendingSelections
    .map((pending, i) => {
      let header = `--- selection ${i + 1} (context from the open document) ---\n`
      if (pending.slide) header += `Slide: ${Array.isArray(pending.slide) ? pending.slide.join(', ') : pending.slide}\n`
      const shapes = pending.shapes || (pending.shape ? [pending.shape] : [])
      if (shapes.length) {
        header += 'Shapes:\n'
        shapes.forEach((s) => {
          header += `  - ${s.name || 'unnamed'} (id ${s.id}, left ${Math.round(s.left)}pt, top ${Math.round(s.top)}pt, ${Math.round(s.width)}x${Math.round(s.height)}pt)\n`
        })
      }
      return `${header}${pending.text}\n--- end of selection ---`
    })
    .join('\n')
}

/* Auto-grow the textarea with its content, capped so the pane keeps its
   proportions; shrink again after sending. */
var INPUT_MAX_HEIGHT = 200

function autoGrowInput() {
  const input = document.getElementById('oc-input')
  if (!input) return
  input.style.height = 'auto'
  input.style.height = `${Math.min(input.scrollHeight, INPUT_MAX_HEIGHT)}px`
  input.style.overflowY = input.scrollHeight > INPUT_MAX_HEIGHT ? 'auto' : 'hidden'
}

/* Local Office.js execution — same AsyncFunction pattern as the bridge's
   executeCode, but the taskpane is the caller, so no WS round-trip. */
function runLocalOffice(code) {
  var host = detectHost()
  if (!host) return Promise.reject(new Error('No Office host available'))
  var AsyncFunction = (async () => {}).constructor
  return host.api.run((context) => {
    var fn = new AsyncFunction('context', host.name, code)
    return fn(context, host.api)
  })
}

/* Rendering */

/* Incremental rendering — only append messages that are not on screen yet.
   A full wipe-and-rebuild replays every bubble's entry animation, which reads
   as flicker on every refetch. */
function renderFromRows(rows) {
  const visible = rows.filter((row) => !OpenCodeChat.isTerminalRow(row)).reverse() // server returns newest-first; render chronologically
  const ids = visible.map((row) => row.id)
  const rendered = chatState.renderedIds || []

  // Prefix match → history is append-only; only render the new tail.
  // Otherwise (session switch, reverted messages) rebuild from scratch.
  const isAppend = ids.length >= rendered.length && rendered.every((id, i) => ids[i] === id)
  const messages = document.getElementById('oc-messages')
  if (!isAppend) {
    messages.innerHTML = ''
    chatState.renderedIds = []
  }

  const streamingEl = messages.querySelector('.chat-message.streaming')
  let streamingReused = false
  for (let i = chatState.renderedIds.length; i < visible.length; i++) {
    const row = visible[i]
    const text = OpenCodeChat.messageText(row)
    if (!text) {
      chatState.renderedIds.push(row.id)
      continue
    }
    const role = row.type === 'user' ? 'user' : 'assistant'
    // The optimistic user bubble (added at send time, before its server id is
    // known) is adopted instead of appended again.
    if (role === 'user') {
      const userBubbles = messages.querySelectorAll('.chat-message.user')
      const last = userBubbles[userBubbles.length - 1]
      if (last && last.dataset.pending === '1' && last.textContent === text) {
        delete last.dataset.pending
        chatState.renderedIds.push(row.id)
        continue
      }
    }
    // A finished assistant message replaces the streaming bubble in place —
    // same content, so reusing the element avoids a visible swap.
    if (streamingEl && !streamingReused && role === 'assistant') {
      streamingEl.className = 'chat-message assistant'
      streamingEl.textContent = text
      streamingReused = true
      chatState.renderedIds.push(row.id)
      continue
    }
    appendMessage(role, text)
    chatState.renderedIds.push(row.id)
  }
  if (!chatState.streamingText) setChatStatus('')
  scrollMessages()
}

function appendMessage(role, text, pending) {
  var messages = document.getElementById('oc-messages')
  var div = document.createElement('div')
  div.className = `chat-message ${role}`
  div.textContent = text
  if (pending) div.dataset.pending = '1'
  messages.appendChild(div)
  scrollMessages()
}

function renderStreaming() {
  var existing = document.querySelector('.chat-message.streaming')
  if (!chatState.streamingText) {
    if (existing) existing.remove()
    return
  }
  if (!existing) {
    existing = document.createElement('div')
    existing.className = 'chat-message assistant streaming'
    document.getElementById('oc-messages').appendChild(existing)
  }
  existing.textContent = chatState.streamingText
  scrollMessages()
}

function scrollMessages() {
  const messages = document.getElementById('oc-messages')
  if (!messages) return
  // rAF: the browser needs a laid-out frame before scrollTop sticks
  requestAnimationFrame(() => {
    messages.scrollTop = messages.scrollHeight
  })
}

function lastAssistantText() {
  var nodes = document.querySelectorAll('#oc-messages .chat-message.assistant:not(.streaming)')
  return nodes.length ? nodes[nodes.length - 1].textContent : ''
}

/* Tab switching — the bridge keeps running in the background either way */
function showMode(mode) {
  document.getElementById('mode-bridge').classList.toggle('hidden', mode !== 'bridge')
  document.getElementById('mode-chat').classList.toggle('hidden', mode !== 'chat')
  document.getElementById('tab-bridge').classList.toggle('active', mode === 'bridge')
  document.getElementById('tab-chat').classList.toggle('active', mode === 'chat')
}

document.getElementById('tab-bridge').addEventListener('click', () => {
  showMode('bridge')
})
document.getElementById('tab-chat').addEventListener('click', () => {
  showMode('chat')
  focusInput()
  scrollMessages()
})
initChat()
