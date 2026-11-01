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
  reasoningRendered: 0, // chars already appended to the strip (incremental render)
  renderedIds: [], // message ids already on screen (incremental rendering)
  inputHistory: [], // sent prompts, newest last (ArrowUp recalls them)
  historyIndex: -1, // -1 = live input, otherwise an index into inputHistory
  draft: '', // what the user typed before browsing history
}

var INPUT_HISTORY_KEY = 'olm-input-history'
var INPUT_HISTORY_LIMIT = 50

function loadInputHistory() {
  try {
    const raw = window.localStorage.getItem(INPUT_HISTORY_KEY)
    const parsed = raw ? JSON.parse(raw) : []
    chatState.inputHistory = Array.isArray(parsed) ? parsed.filter((x) => typeof x === 'string') : []
  } catch {
    chatState.inputHistory = [] // private mode or corrupt storage — start empty
  }
}

function rememberInput(text) {
  const trimmed = text.trim()
  if (!trimmed) return
  // Consecutive duplicates occupy one history slot.
  if (chatState.inputHistory[chatState.inputHistory.length - 1] === trimmed) return
  chatState.inputHistory.push(trimmed)
  if (chatState.inputHistory.length > INPUT_HISTORY_LIMIT) {
    chatState.inputHistory = chatState.inputHistory.slice(-INPUT_HISTORY_LIMIT)
  }
  try {
    window.localStorage.setItem(INPUT_HISTORY_KEY, JSON.stringify(chatState.inputHistory))
  } catch {
    // Storage unavailable — the in-memory history still works for this panel.
  }
}

/* ArrowUp / ArrowDown walk the sent prompts, newest first. Browsing keeps the
   user's draft; going past the newest entry restores it (terminal-style). */
function historyMove(delta) {
  const input = document.getElementById('oc-input')
  if (!chatState.inputHistory.length) return
  if (chatState.historyIndex === -1) {
    if (delta > 0) return // already at the live input, nothing newer
    chatState.draft = input.value
    chatState.historyIndex = chatState.inputHistory.length - 1
  } else {
    chatState.historyIndex += delta
    if (chatState.historyIndex >= chatState.inputHistory.length) {
      chatState.historyIndex = -1
      input.value = chatState.draft
      input.setSelectionRange(chatState.draft.length, chatState.draft.length)
      autoGrowInput()
      return
    }
  }
  if (chatState.historyIndex < 0) chatState.historyIndex = 0
  const value = chatState.inputHistory[chatState.historyIndex]
  input.value = value
  input.setSelectionRange(value.length, value.length)
  autoGrowInput()
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
  // Connection config comes from localStorage (loadConfig); the pane itself
  // offers no connection UI — the same-origin /opencode proxy is the default.
  chatState.config = OpenCodeChat.loadConfig()

  document.getElementById('oc-new-session').addEventListener('click', () => {
    OpenCodeChat.createSession(chatState.config, 'Office add-in chat').then((session) => {
      selectSession(session.id)
      refreshSessions()
      focusInput()
    }, chatError)
  })

  document.getElementById('oc-send').addEventListener('click', sendChatMessage)
  document.getElementById('oc-session').addEventListener('change', (event) => {
    if (event.target.value) selectSession(event.target.value)
  })
  loadInputHistory()
  const input = document.getElementById('oc-input')
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      sendChatMessage()
    } else if (event.key === 'ArrowUp' && !event.shiftKey && !event.altKey && !event.metaKey && !event.ctrlKey) {
      event.preventDefault()
      historyMove(-1)
    } else if (event.key === 'ArrowDown' && !event.shiftKey && !event.altKey && !event.metaKey && !event.ctrlKey) {
      event.preventDefault()
      historyMove(1)
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
      let text = String(out?.text ?? '')
        .replace(/\r/g, '')
        .replace(/\n{3,}/g, '\n\n')
        .trim()
      const shapes = out?.shapes ?? []
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
    // Most recent activity first: the API order is not guaranteed.
    var updatedAt = (session) => session?.time?.updated ?? 0
    sessions
      .slice()
      .sort((a, b) => updatedAt(b) - updatedAt(a))
      .slice(0, 25)
      .forEach((session) => {
        var option = document.createElement('option')
        option.value = session.id
        option.textContent = session.title || session.id
        select.appendChild(option)
      })
    if (chatState.sessionID) select.value = chatState.sessionID
    // Auto-select only when no session is active yet; re-selecting the
    // current session on every refresh would wipe and rebuild the bubbles.
    if (select.value && !chatState.sessionID) selectSession(select.value)
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
    } else if (event.type === 'session.reasoning.started') {
      chatState.reasoningActive = true
      updateReasoningStrip()
    } else if (event.type === 'session.reasoning.delta') {
      // A delta without a preceding started/step event still means live reasoning.
      chatState.reasoningActive = true
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
  rememberInput(typed)
  chatState.historyIndex = -1
  chatState.draft = ''
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

/* Reasoning strip — a news-ticker above the messages: while the assistant
   works, the reasoning crawls by in one unreadable line (two copies of the
   text, translated -50% in a loop, so the band never shows a gap). After
   the step it collapses to a label-only summary. Rendering is incremental:
   each delta only appends its own text nodes, never a full rewrite. */
/* Realtime tail-follow ticker: the newest reasoning stays visible at the
   right edge and the band moves left exactly as fast as reasoning text
   arrives — the speed IS the streaming rate. A px/s cap keeps delta bursts
   from teleporting the band; when no deltas arrive, the band stands still.
   The band follows the tail of the FIRST copy only: pinning the end of
   copy 2 would make it travel double the distance and lag a full
   copy-width behind. */
var TICKER_MAX_SPEED = 600 // px/s
var tickerRaf = 0
var tickerOffset = null // null = snap to the tail on the next frame
var tickerLast = 0

function tickTicker(ts) {
  var strip = document.getElementById('oc-reasoning-strip')
  var body = document.getElementById('oc-reasoning-text')
  if (!strip || !body || !strip.classList.contains('live')) {
    tickerRaf = 0
    return
  }
  var style = window.getComputedStyle(strip)
  var viewport = strip.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight)
  var copyWidth = body.children[0] ? body.children[0].getBoundingClientRect().width : 0
  var target = Math.min(0, viewport - copyWidth)
  var dt
  var want
  if (tickerOffset === null || target > tickerOffset) {
    tickerOffset = target // fresh text (new step) or text that shrank: snap
  } else if (tickerLast) {
    dt = (ts - tickerLast) / 1000
    want = tickerOffset - target
    tickerOffset -= Math.min(want, TICKER_MAX_SPEED * dt)
  }
  body.style.transform = `translateX(${tickerOffset}px)`
  tickerLast = ts
  tickerRaf = window.requestAnimationFrame(tickTicker)
}

function startTicker() {
  var body = document.getElementById('oc-reasoning-text')
  if (!body) return
  if (!tickerRaf) {
    tickerLast = 0
    tickerRaf = window.requestAnimationFrame(tickTicker)
  }
}

function stopTicker() {
  if (tickerRaf) window.cancelAnimationFrame(tickerRaf)
  tickerRaf = 0
  tickerLast = 0
  var body = document.getElementById('oc-reasoning-text')
  if (body) body.style.transform = ''
  tickerOffset = null
}

function updateReasoningStrip() {
  var strip = document.getElementById('oc-reasoning-strip')
  if (!strip) return
  var text = chatState.reasoningText
  var body = document.getElementById('oc-reasoning-text')
  var copies = body ? body.children : []
  var label = null
  var i
  var delta
  if (!text) {
    strip.hidden = true
    strip.dataset.expanded = ''
    chatState.reasoningRendered = 0
    tickerOffset = null
    if (body) {
      for (i = 0; i < copies.length; i++) copies[i].textContent = ''
    }
    return
  }
  strip.hidden = false
  label = document.getElementById('oc-reasoning-label')
  var rendered = chatState.reasoningRendered || 0
  if (text.length >= rendered && text.slice(0, rendered) === (copies[0] ? copies[0].textContent : '')) {
    // Normal case: append only the new tail, to both ticker copies.
    if (text.length > rendered) {
      delta = document.createTextNode(text.slice(rendered))
      for (i = 0; i < copies.length; i++) copies[i].appendChild(delta.cloneNode(false))
    }
  } else {
    // Reset (new step): full rewrite is fine here, it happens once.
    for (i = 0; i < copies.length; i++) copies[i].textContent = text
    tickerOffset = null // new text: snap the band to its tail on the next frame
  }
  chatState.reasoningRendered = text.length
  var expanded = false
  if (chatState.reasoningActive) {
    strip.classList.add('live')
    label.textContent = 'Reasoning…'
    body.classList.remove('expanded')
    startTicker()
  } else {
    strip.classList.remove('live')
    stopTicker()
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
   as flicker on every refetch. Session switches load the full history. */
function windowShiftIndex(ids, rendered) {
  if (!rendered.length) return -1
  for (let j = 0; j + rendered.length <= ids.length; j++) {
    let ok = true
    for (let k = 0; k < rendered.length; k++) {
      if (ids[j + k] !== rendered[k]) {
        ok = false
        break
      }
    }
    if (ok) return j
  }
  return -1
}

function renderFromRows(rows) {
  const visible = rows.filter((row) => !OpenCodeChat.isTerminalRow(row)).reverse() // server returns newest-first; render chronologically
  const ids = visible.map((row) => row.id)
  const rendered = chatState.renderedIds || []

  // The old window reappears shifted inside the new one → render only the
  // new tail and drop the messages that slid out of the window. Otherwise
  // (session switch, reverted messages) rebuild from scratch.
  const shift = windowShiftIndex(ids, rendered)
  const isAppend = shift >= 0
  const messages = document.getElementById('oc-messages')
  if (!isAppend) {
    messages.innerHTML = ''
    chatState.renderedIds = []
  } else if (shift > 0) {
    for (let d = 0; d < shift; d++) {
      const first = messages.querySelector('.chat-message')
      if (first) first.remove()
    }
  }

  const streamingEl = messages.querySelector('.chat-message.streaming')
  let streamingReused = false
  const firstNew = isAppend ? shift + rendered.length : 0
  for (let i = firstNew; i < visible.length; i++) {
    const row = visible[i]
    const text = OpenCodeChat.messageText(row)
    const images = OpenCodeChat.messageImages(row)
    if (!text && !images.length) {
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
      fillMessageBubble(streamingEl, text, images)
      streamingReused = true
      chatState.renderedIds.push(row.id)
      continue
    }
    appendMessage(role, text, false, images)
    chatState.renderedIds.push(row.id)
  }
  if (!chatState.streamingText) setChatStatus('')
  scrollMessages(true)
}

function appendMessage(role, text, pending, images) {
  var messages = document.getElementById('oc-messages')
  var div = document.createElement('div')
  div.className = `chat-message ${role}`
  fillMessageBubble(div, text, images)
  if (pending) div.dataset.pending = '1'
  messages.appendChild(div)
  scrollMessages(true)
}

/* Text plus any model-rendered images (file parts) inside one bubble. */
function fillMessageBubble(div, text, images) {
  div.textContent = text
  ;(images || []).forEach((image) => {
    var img = document.createElement('img')
    img.className = 'chat-image'
    img.src = image.src
    img.alt = image.name
    div.appendChild(img)
  })
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

function scrollMessages(smooth) {
  const messages = document.getElementById('oc-messages')
  if (!messages) return
  // rAF: the browser needs a laid-out frame before scrollTop sticks. History
  // loads glide to the bottom; streaming deltas follow instantly so the
  // smooth animation never lags behind the text.
  requestAnimationFrame(() => {
    messages.scrollTo({ top: messages.scrollHeight, behavior: smooth ? 'smooth' : 'auto' })
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
