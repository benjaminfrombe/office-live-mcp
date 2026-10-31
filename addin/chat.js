/* OpenCode chat UI — taskpane mode that chats with the OpenCode server and
   applies replies to the live document via Office.js. Runs next to the bridge
   mode in the same taskpane; the bridge keeps working while the chat is open. */

var chatState = {
  config: null,
  sessionID: null,
  events: null, // active SSE subscription
  streamingText: '', // partial assistant text while streaming
  busy: false,
}

var EVENT_TYPES = [
  'session.step.started',
  'session.text.delta',
  'session.text.ended',
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

  document.getElementById('oc-use-selection').addEventListener('click', () => {
    runLocalOffice(
      'const selection = await new Promise((resolve) => ' +
        'Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => resolve(r.value))); ' +
        'return String(selection || "")',
    ).then((text) => {
      const input = document.getElementById('oc-input')
      if (text) {
        input.value = `${input.value ? `${input.value}\n` : ''}--- selected text ---\n${text}\n--- end ---`
      }
      focusInput()
      // caret at the very end, past the pasted selection
      input.setSelectionRange(input.value.length, input.value.length)
      autoGrowInput()
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
}

function setChatStatus(text) {
  var el = document.getElementById('oc-status')
  if (el) el.textContent = text
}

function chatError(err) {
  console.error('OpenCode chat error:', err)
  setChatStatus(String(err?.message ? err.message : err))
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
  document.getElementById('oc-session').value = sessionID
  renderFromRows([])
  setChatStatus('Loading messages...')
  OpenCodeChat.getMessages(chatState.config, sessionID).then(renderFromRows, chatError)
  startEventStream()
}

/* Event stream — deltas append to streamingText; terminal events trigger a
   full message refetch so the rendered history always matches the server. */
function startEventStream() {
  var sessionID = chatState.sessionID
  if (!sessionID) return
  chatState.events = OpenCodeChat.subscribeEvents(chatState.config, EVENT_TYPES, (event) => {
    if (event.type === 'stream.closed' || event.type === 'stream.error') {
      if (chatState.sessionID === sessionID) {
        setChatStatus(event.type === 'stream.error' ? 'Event stream error' : '')
        setTimeout(startEventStream, 2000)
      }
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
    } else if (event.type === 'session.step.started') {
      setChatStatus('Assistant is working...')
    } else if (
      event.type === 'session.step.ended' ||
      event.type === 'session.step.failed' ||
      event.type === 'session.execution.succeeded' ||
      event.type === 'session.execution.failed' ||
      event.type === 'session.execution.interrupted'
    ) {
      chatState.streamingText = ''
      setChatStatus('')
      OpenCodeChat.getMessages(chatState.config, sessionID).then(renderFromRows, chatError)
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
  const text = input.value.trim()
  if (!text || !chatState.sessionID || chatState.busy) {
    focusInput()
    return
  }
  chatState.busy = true
  input.value = ''
  setChatStatus('Sending...')
  appendMessage('user', text)
  autoGrowInput()
  focusInput()
  OpenCodeChat.sendPrompt(chatState.config, chatState.sessionID, text).then(
    () => {
      chatState.busy = false
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

function renderFromRows(rows) {
  const messages = document.getElementById('oc-messages')
  messages.innerHTML = ''
  rows
    .filter((row) => !OpenCodeChat.isTerminalRow(row))
    .reverse() // server returns newest-first; render chronologically
    .forEach((row) => {
      var role = row.type === 'user' ? 'user' : 'assistant'
      var text = OpenCodeChat.messageText(row)
      if (text) appendMessage(role, text)
    })
  if (!chatState.streamingText) setChatStatus('')
  scrollMessages()
}

function appendMessage(role, text) {
  var messages = document.getElementById('oc-messages')
  var div = document.createElement('div')
  div.className = `chat-message ${role}`
  div.textContent = text
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
