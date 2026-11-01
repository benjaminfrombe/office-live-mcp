// Reasoning-ticker mockup: simulates the OpenCode session event flow with
// mock data and drives the same DOM the real addin uses, so CSS fixes made
// in addin/style.css can be evaluated here without a live session.
(function () {
  var state = { reasoningText: '', reasoningActive: false, reasoningRendered: 0 }
  var variant = 'ticker'
  var maxFollowSpeed = 600 // px/s cap — well above a normal stream so the band
  // tracks the tail at the real delta rate; only bursts get capped
  var timers = []
  var ticker = { raf: 0, offset: 0, last: 0 }

  function el(id) { return document.getElementById(id) }

  /* Realtime tail-follow ticker: the newest reasoning stays visible at the
     right edge and the band moves left exactly as fast as text arrives —
     the speed IS the streaming rate. A px/s cap keeps delta bursts from
     teleporting the band; when no deltas arrive, the band stands still. */
  function tickTicker(ts) {
    var body = el('oc-reasoning-text')
    if (!body || variant !== 'ticker' || !el('oc-reasoning-strip').classList.contains('live')) {
      ticker.raf = 0
      return
    }
    var strip = el('oc-reasoning-strip')
    // The text element is width:max-content, so its own clientWidth never
    // overflows; the viewport is the strip's content box. Follow the tail of
    // the FIRST copy only — pinning the end of copy 2 would make the band
    // travel double the distance and lag a full copy-width behind.
    var cs = getComputedStyle(strip)
    var viewport = strip.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)
    var copyWidth = body.children[0] ? body.children[0].getBoundingClientRect().width : 0
    var target = viewport - copyWidth
    if (target > 0) target = 0
    if (ticker.offset === null || target > ticker.offset) {
      // Fresh text (new step) or text that shrank: snap to the tail.
      ticker.offset = target
    } else if (ticker.last) {
      var dt = (ts - ticker.last) / 1000
      var want = ticker.offset - target // positive distance still to travel
      var step = Math.min(want, maxFollowSpeed * dt)
      ticker.offset -= step
    }
    body.style.transform = 'translateX(' + ticker.offset + 'px)'
    ticker.last = ts
    ticker.raf = requestAnimationFrame(tickTicker)
  }

  function startTicker() {
    var body = el('oc-reasoning-text')
    if (!body) return
    body.style.animation = 'none'
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

  // ---- reasoning strip (copy of the real incremental renderer) ----
  function updateReasoningStrip() {
    var strip = el('oc-reasoning-strip')
    var text = state.reasoningText
    var body = el('oc-reasoning-text')
    var copies = body.children
    if (!text) {
      strip.hidden = true
      strip.dataset.expanded = ''
      state.reasoningRendered = 0
      for (var i = 0; i < copies.length; i++) copies[i].textContent = ''
      return
    }
    strip.hidden = false
    var label = el('oc-reasoning-label')
    var rendered = state.reasoningRendered || 0
    if (text.length >= rendered && text.slice(0, rendered) === (copies[0] ? copies[0].textContent : '')) {
      if (text.length > rendered) {
        var delta = document.createTextNode(text.slice(rendered))
        for (var j = 0; j < copies.length; j++) copies[j].appendChild(delta.cloneNode(false))
      }
    } else {
      // Reset (new step): full rewrite is fine here, it happens once.
      for (var k = 0; k < copies.length; k++) copies[k].textContent = text
      ticker.offset = null // new text: snap the band to its tail on the next frame
    }
    state.reasoningRendered = text.length
    if (state.reasoningActive) {
      strip.classList.add('live')
      label.textContent = 'Reasoning…'
      body.classList.remove('expanded')
      startTicker()
    } else {
      strip.classList.remove('live')
      body.style.transform = ''
      var expanded = strip.dataset.expanded === '1'
      label.textContent = 'Reasoning · ' + text.length + ' chars · click to ' + (expanded ? 'collapse' : 'expand')
      body.classList.toggle('expanded', expanded)
      if (expanded) body.scrollTop = body.scrollHeight
    }
  }

  el('oc-reasoning-strip').addEventListener('click', function () {
    var strip = el('oc-reasoning-strip')
    if (strip.classList.contains('live')) return
    strip.dataset.expanded = strip.dataset.expanded === '1' ? '' : '1'
    updateReasoningStrip()
  })

  // ---- variant plumbing (mockup-only CSS overrides) ----
  function applyVariant() {
    var body = el('oc-reasoning-text')
    var style = el('mock-variant-style')
    if (!style) {
      style = document.createElement('style')
      style.id = 'mock-variant-style'
      document.head.appendChild(style)
    }
    body.classList.remove('variant-tail', 'variant-fade')
    var css = ''
    if (variant === 'tail') {
      // B: no crawl — the box follows the tail like a terminal, one line tall.
      css = '.reasoning-strip.live .reasoning-text{animation:none;white-space:pre-wrap;display:block;width:auto;' +
        'max-height:1.5em;overflow:hidden;}' +
        '.reasoning-strip.live .reasoning-text>span::after{content:"";}' +
        '.reasoning-strip.live .reasoning-text.variant-tail>span:last-child{display:block;transform:translateY(-50%);}'
      body.classList.add('variant-tail')
    } else if (variant === 'fade') {
      // C: static single line, newest characters highlighted as they arrive.
      css = '.reasoning-strip.live .reasoning-text{animation:none;white-space:nowrap;display:block;width:auto;' +
        'max-height:1.5em;overflow:hidden;direction:rtl;text-align:left;}' +
        '.reasoning-strip.live .reasoning-text>span::after{content:"";}'
      body.classList.add('variant-fade')
    } else {
      css = '' // variant A: JS rAF ticker drives the transform, no CSS animation
    }
    style.textContent = css
    if (variant === 'ticker') startTicker()
    else stopTicker()
  }

  el('mock-variant').addEventListener('change', function (e) {
    variant = e.target.value
    applyVariant()
    // Re-render from scratch so the variant applies to full text.
    state.reasoningRendered = 0
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
    ticker.offset = 0
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
    timers.push(setTimeout(function () {
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
                // The model rendered a slide check image: file part in the reply.
                fillMessageBubble(bubble, bubble.textContent, [{ src: MOCK_SLIDE_CHECK, name: 'slide-1-check.png' }])
              })
            }, 700))
          })
        }, 900))
      })
    }, 500))
  }

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
  applyVariant()
  runConversation('Make the title on slide 1 bigger.')
})()
