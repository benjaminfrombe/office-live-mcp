// Reasoning-ticker mockup: simulates the OpenCode session event flow with
// mock data and drives the same DOM the real addin uses, so CSS fixes made
// in addin/style.css can be evaluated here without a live session.
(function () {
  var state = { reasoningText: '', reasoningActive: false, reasoningRendered: 0 }
  var maxFollowSpeed = 600 // px/s cap — well above a normal stream so the band
  // tracks the tail at the real delta rate; only bursts get capped
  var timers = []
  var ticker = { raf: 0, offset: null, last: 0 }

  function el(id) { return document.getElementById(id) }

  /* Realtime tail-follow ticker — mirrors tickTicker in addin/chat.js: the
     newest reasoning stays at the right edge of .reasoning-window and the
     band moves left as fast as text arrives (capped at maxFollowSpeed). */
  function tickTicker(ts) {
    var strip = el('oc-reasoning-strip')
    var body = el('oc-reasoning-text')
    if (!body || !strip.classList.contains('live') || strip.classList.contains('expanded')) {
      ticker.raf = 0
      return
    }
    var viewport = body.parentElement.clientWidth // .reasoning-window
    var copyWidth = body.children[0] ? body.children[0].getBoundingClientRect().width : 0
    var target = Math.min(0, viewport - copyWidth)
    if (ticker.offset === null || target > ticker.offset) {
      ticker.offset = target
    } else if (ticker.last) {
      var dt = (ts - ticker.last) / 1000
      ticker.offset -= Math.min(ticker.offset - target, maxFollowSpeed * dt)
    }
    body.style.transform = 'translateX(' + ticker.offset + 'px)'
    ticker.last = ts
    if (ticker.offset === target) {
      ticker.raf = 0
      ticker.last = 0
      return
    }
    ticker.raf = requestAnimationFrame(tickTicker)
  }

  function startTicker() {
    if (!ticker.raf) {
      ticker.last = 0
      ticker.raf = requestAnimationFrame(tickTicker)
    }
  }

  function stopTicker() {
    if (ticker.raf) cancelAnimationFrame(ticker.raf)
    ticker.raf = 0
    ticker.last = 0
    var body = el('oc-reasoning-text')
    if (body) body.style.transform = ''
    ticker.offset = null
  }

  // ---- message rendering (minimal, mirrors the real classes) ----
  function addMessage(role, text) {
    var box = el('oc-messages')
    var msg = document.createElement('div')
    msg.className = 'chat-message ' + role
    fillMessageBubble(msg, text, null)
    box.appendChild(msg)
    box.scrollTop = box.scrollHeight
    return msg
  }

  /* Mirrors the real fillMessageBubble: text plus model-rendered images. */
  function fillMessageBubble(div, text, images) {
    div.textContent = text
    ;(images || []).forEach(function (image) {
      var img = document.createElement('img')
      img.className = 'chat-image'
      img.src = image.src
      img.alt = image.name
      div.appendChild(img)
    })
  }

  // Mock "rendered slide" the LLM allegedly produced to check its own work.
  var MOCK_SLIDE_CHECK = 'data:image/svg+xml;utf8,' + encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180">' +
      '<rect width="320" height="180" fill="#ffffff"/>' +
      '<rect width="320" height="34" fill="#0b5e0b"/>' +
      '<text x="16" y="22" font-family="Helvetica" font-size="14" fill="#ffffff">Slide 1 — checked render</text>' +
      '<text x="16" y="80" font-family="Helvetica" font-size="22" font-weight="bold" fill="#26302a">delaware</text>' +
      '<text x="16" y="104" font-family="Helvetica" font-size="12" fill="#6b756e">Title scaled to 50pt, inside safe area</text>' +
      '<rect x="16" y="124" width="120" height="8" rx="4" fill="#107c10"/>' +
      '<rect x="16" y="140" width="80" height="8" rx="4" fill="#c9d6c9"/>' +
      '<text x="230" y="168" font-family="Helvetica" font-size="9" fill="#9aa79a">mockup image · file part</text>' +
      '</svg>'
  )

  function streamInto(node, text, cps, done) {
    var i = 0
    function tick() {
      var chunk = text.slice(i, i + 3)
      node.textContent += chunk
      i += chunk.length
      var box = el('oc-messages')
      box.scrollTop = box.scrollHeight
      if (i < text.length) timers.push(setTimeout(tick, 1000 / cps))
      else if (done) done(node)
    }
    tick()
  }

  // ---- reasoning strip (mirrors updateReasoningStrip in addin/chat.js) ----
  function updateReasoningStrip() {
    var strip = el('oc-reasoning-strip')
    var text = state.reasoningText
    var body = el('oc-reasoning-text')
    var copies = body.children
    if (!text) {
      strip.hidden = true
      strip.dataset.expanded = ''
      strip.classList.remove('expanded')
      state.reasoningRendered = 0
      ticker.offset = null
      for (var i = 0; i < copies.length; i++) copies[i].textContent = ''
      return
    }
    strip.hidden = false
    var rendered = state.reasoningRendered || 0
    if (text.length >= rendered && text.slice(0, rendered) === (copies[0] ? copies[0].textContent : '')) {
      if (text.length > rendered) {
        var delta = document.createTextNode(text.slice(rendered))
        for (var j = 0; j < copies.length; j++) copies[j].appendChild(delta.cloneNode(false))
      }
    } else {
      for (var k = 0; k < copies.length; k++) copies[k].textContent = text
      ticker.offset = null
    }
    state.reasoningRendered = text.length
    var expanded = strip.dataset.expanded === '1'
    strip.classList.toggle('live', state.reasoningActive)
    strip.classList.toggle('expanded', expanded)
    strip.title = expanded ? 'Click to collapse' : 'Click to expand (' + text.length + ' chars)'
    if (expanded) {
      stopTicker()
      body.scrollTop = body.scrollHeight
    } else if (state.reasoningActive) {
      startTicker()
    } else {
      stopTicker()
    }
  }

  el('oc-reasoning-strip').addEventListener('click', function (e) {
    var strip = el('oc-reasoning-strip')
    if (strip.classList.contains('expanded') && e.target.closest('.reasoning-window')) return
    strip.dataset.expanded = strip.dataset.expanded === '1' ? '' : '1'
    updateReasoningStrip()
  })

  el('mock-speed').addEventListener('input', function (e) {
    maxFollowSpeed = Number(e.target.value)
    el('mock-speed-value').textContent = maxFollowSpeed + ' px/s'
  })

  // ---- scripted conversation (mock data) ----
  var REASONING_1 = [
    'The user wants the title on slide 1 larger. ',
    'First I need the current shape geometry, then I can scale the text ',
    'without pushing the title out of the safe area. ',
    'Slide width is 10in, the title box spans most of it, so a 1.25x font ',
    'scale should fit. I will read the shape first to confirm the name. ',
  ].join('')
  var REASONING_2 = [
    'The read returned "Title 1" at 40pt. Scaling to 50pt keeps the box ',
    'within bounds. I will set the text size and verify the resulting ',
    'geometry does not overlap the subtitle placeholder. ',
  ].join('')

  function reasoningStream(text, cps, done) {
    var i = 0
    function tick() {
      state.reasoningText += text.slice(i, i + 4)
      i += 4
      updateReasoningStrip()
      if (i < text.length) timers.push(setTimeout(tick, 1000 / cps))
      else if (done) done()
    }
    tick()
  }

  // Mirrors setChatStatus in addin/chat.js (busy / info / error).
  var statusTimer = null
  function setStatus(text, kind) {
    var node = el('oc-status')
    node.textContent = text
    node.dataset.kind = text ? kind || 'info' : ''
    clearTimeout(statusTimer)
    if (text && (kind || 'info') === 'info') statusTimer = setTimeout(function () { setStatus('') }, 4000)
  }

  function clearTimers() {
    timers.forEach(clearTimeout)
    timers = []
  }

  function reset() {
    clearTimers()
    el('oc-messages').textContent = ''
    state.reasoningText = ''
    state.reasoningActive = false
    state.reasoningRendered = 0
    stopTicker()
    updateReasoningStrip()
  }

  // ---- input history (ArrowUp / ArrowDown through sent messages) ----
  var history = [
    'Make the title on slide 1 bigger.',
    'Remove the empty slide 4.',
    'Limit the bullets on slide 2 to three lines.',
  ]
  var historyIndex = -1 // -1 = live input (not browsing history)
  var draft = '' // what the user typed before browsing history

  function remember(text) {
    var trimmed = text.trim()
    if (!trimmed) return
    // Consecutive duplicates occupy one history slot.
    if (history[history.length - 1] === trimmed) return
    history.push(trimmed)
  }

  function historyMove(delta) {
    var input = document.querySelector('.chat-input textarea')
    if (!history.length) return
    if (historyIndex === -1) {
      if (delta > 0) return // already at the live input, nothing newer
      draft = input.value
      historyIndex = history.length - 1
    } else {
      historyIndex += delta
      if (historyIndex >= history.length) {
        // Past the newest entry: back to the live input with the draft.
        historyIndex = -1
        input.value = draft
        input.setSelectionRange(draft.length, draft.length)
        return
      }
    }
    if (historyIndex < 0) historyIndex = 0
    var value = history[historyIndex]
    input.value = value
    input.setSelectionRange(value.length, value.length)
  }

  function send() {
    var input = document.querySelector('.chat-input textarea')
    var text = input.value.trim()
    if (!text) return
    remember(text)
    historyIndex = -1
    draft = ''
    input.value = ''
    runConversation(text)
  }

  document.querySelector('.chat-input textarea').addEventListener('keydown', function (e) {
    if (e.key === 'ArrowUp' && !e.shiftKey && !e.altKey && !e.metaKey && !e.ctrlKey) {
      e.preventDefault()
      historyMove(-1)
    } else if (e.key === 'ArrowDown' && !e.shiftKey && !e.altKey && !e.metaKey && !e.ctrlKey) {
      e.preventDefault()
      historyMove(1)
    } else if (e.key === 'Enter' && !e.shiftKey && !e.metaKey && !e.ctrlKey) {
      e.preventDefault()
      send()
    }
  })
  document.querySelector('.chat-input .primary').addEventListener('click', send)

  function runConversation(question) {
    reset()
    addMessage('user', question)
    setStatus('Sending…', 'busy')
    timers.push(setTimeout(function () {
      setStatus('Working…', 'busy')
      state.reasoningActive = true
      reasoningStream(REASONING_1, 40, function () {
        timers.push(setTimeout(function () {
          // Step boundary: reasoning resets, second step streams.
          state.reasoningText = ''
          state.reasoningRendered = 0
          updateReasoningStrip()
          reasoningStream(REASONING_2, 40, function () {
            timers.push(setTimeout(function () {
              state.reasoningActive = false
              updateReasoningStrip()
              streamInto(addMessage('assistant', ''), 'Done — the title on slide 1 is now 50pt and stays inside the safe area.', 60, function (bubble) {
                setStatus('')
                // The model rendered a slide check image: file part in the reply.
                fillMessageBubble(bubble, bubble.textContent, [{ src: MOCK_SLIDE_CHECK, name: 'slide-1-check.png' }])
              })
            }, 700))
          })
        }, 900))
      })
    }, 500))
  }

  // ---- spacing checks: a question form card and a selection chip, built
  // with the same classes as addin/forms.js and renderSelectionChips ----
  function node(tag, className, text) {
    var n = document.createElement(tag)
    if (className) n.className = className
    if (text) n.textContent = text
    return n
  }

  el('mock-form').addEventListener('click', function () {
    var box = el('oc-form-cards')
    if (box.firstChild) {
      box.textContent = ''
      return
    }
    var card = node('div', 'form-card')
    var head = node('div', 'form-card-head')
    head.appendChild(node('span', 'form-card-title', 'Which title size?'))
    head.appendChild(node('span', 'form-card-session', 'ses_mock'))
    head.appendChild(node('button', 'form-card-dismiss', '×'))
    card.appendChild(head)
    var field = node('div', 'form-field')
    ;['44pt', '50pt', '56pt'].forEach(function (size, i) {
      var row = node('label', 'form-option')
      var input = node('input')
      input.type = 'radio'
      input.name = 'mock-size'
      input.checked = i === 1
      row.appendChild(input)
      var text = node('span', 'form-option-text')
      text.appendChild(node('span', 'form-option-label', size))
      row.appendChild(text)
      field.appendChild(row)
    })
    card.appendChild(field)
    var foot = node('div', 'form-card-foot')
    foot.appendChild(node('span', 'form-card-status'))
    foot.appendChild(node('button', 'form-card-submit primary', 'Submit'))
    card.appendChild(foot)
    box.appendChild(card)
  })

  el('mock-chip').addEventListener('click', function () {
    var box = el('oc-selection-chips')
    if (box.firstChild) {
      box.textContent = ''
      return
    }
    var chip = node('div', 'selection-chip')
    chip.appendChild(node('span', '', 'Selection · slide 1 · Title 1 · 24 chars · delaware'))
    chip.appendChild(node('button', '', '✕'))
    box.appendChild(chip)
  })

  el('mock-status-info').addEventListener('click', function () { setStatus('Model: claude-opus-5-5') })
  el('mock-status-error').addEventListener('click', function () {
    setStatus('OpenCode not reachable: is the server running?', 'error')
  })

  el('mock-replay').addEventListener('click', function () { runConversation('Make the title on slide 1 bigger.') })
  // Debug hook for automated checks: jump straight to the collapsed state.
  window.__mock = {
    end: function () {
      clearTimers()
      state.reasoningActive = false
      updateReasoningStrip()
    },
    // Deterministic test hook: append text synchronously as if a delta arrived.
    feed: function (extra) {
      clearTimers()
      state.reasoningActive = true
      state.reasoningText += extra
      updateReasoningStrip()
    },
    // Deterministic test hook: start a fresh step with the given text.
    newStep: function (text) {
      clearTimers()
      state.reasoningActive = true
      state.reasoningText = text
      state.reasoningRendered = 0
      updateReasoningStrip()
    },
    // Deterministic test hook: advance the ticker manually (hidden-tab testing).
    tick: function (ts) {
      tickTicker(ts)
    },
    // Show the assistant bubble with the mock slide-check image immediately.
    image: function () {
      clearTimers()
      var bubble = addMessage('assistant', 'Done — the title on slide 1 is now 50pt and stays inside the safe area.')
      fillMessageBubble(bubble, bubble.textContent, [{ src: MOCK_SLIDE_CHECK, name: 'slide-1-check.png' }])
    },
    // Simulate a tool-output image: the tool wrote a file to disk and its
    // output text mentioned the path; the add-in renders it via the bridge's
    // /file endpoint (real route on https://localhost:8791).
    toolImage: function () {
      clearTimers()
      var bubble = addMessage('assistant', 'Rendered the slide to check the layout:')
      fillMessageBubble(bubble, bubble.textContent, [
        { src: 'https://localhost:8791/file?path=%2Ftmp%2Folm-mock-slide-check.svg', name: 'slide-1-render.svg' },
      ])
    },
  }
  runConversation('Make the title on slide 1 bigger.')
})()
