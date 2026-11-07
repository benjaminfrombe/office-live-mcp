/* OpenCode form cards — pending question prompts rendered inside the taskpane.
   The `question` tool surfaces on the OpenCode server as a Form (frm_…); this
   module polls the global form list, renders one card per pending form with
   clickable answer controls, and posts the answer back through the same-origin
   /opencode proxy. Plain JS, no build step — loaded with a script tag. */

var OpenCodeForms = (() => {
  var POLL_MS = 3000

  /* ---------- Pure helpers (unit-tested, no DOM) ---------- */

  /* A field is visible when it is not hidden and every `when` condition
     matches the answers collected so far. A `when` on an unanswered
     dependency keeps the field hidden. */
  function visibleFields(fields, answers) {
    return (fields || []).filter((field) => {
      if (field.hidden) return false
      const when = field.when || []
      return when.every((cond) => {
        const current = answers[cond.key]
        if (current === undefined) return false
        return cond.op === 'neq' ? current !== cond.value : current === cond.value
      })
    })
  }

  function toValue(field, raw) {
    if (field.type === 'integer' || field.type === 'number') {
      const n = Number(raw)
      return raw === '' || raw === undefined || Number.isNaN(n) ? undefined : n
    }
    if (field.type === 'boolean') return raw === true
    return raw
  }

  /* Form.Answer = { [fieldKey]: Form.Value }; unanswered optional fields are
     omitted so the server-side defaults stay in effect. */
  function buildAnswer(fields, state) {
    const values = state?.values || {}
    const custom = state?.custom || {}
    const answer = {}
    for (const field of fields || []) {
      if (field.hidden) continue
      if (field.type === 'multiselect') {
        const list = (values[field.key] || []).slice()
        const extra = (custom[field.key] || '').trim()
        if (field.custom && extra) list.push(extra)
        if (list.length) answer[field.key] = list
        continue
      }
      const value = toValue(field, values[field.key])
      if (value !== undefined && value !== '') answer[field.key] = value
    }
    return answer
  }

  /* Returns an error message for the first invalid visible field, or null. */
  function validate(fields, state) {
    const _values = state?.values || {}
    const custom = state?.custom || {}
    const answer = buildAnswer(fields, state)
    for (const field of fields || []) {
      if (field.hidden) continue
      const hasWhen = (field.when || []).length > 0
      if (hasWhen && !(field.key in answer)) continue // hidden by `when`
      if (field.type === 'multiselect') {
        const count = (answer[field.key] || []).length
        const min = field.required ? Math.max(1, field.minItems || 0) : field.minItems
        if (min !== undefined && count < min) return `${field.title || field.key}: pick at least ${min}`
        if (field.maxItems !== undefined && count > field.maxItems)
          return `${field.title || field.key}: pick at most ${field.maxItems}`
        continue
      }
      if (!field.required) continue
      const value = answer[field.key]
      if (value === undefined || value === '') {
        return `${field.title || field.key}: answer required`
      }
    }
    void custom
    return null
  }

  /* ---------- DOM glue ---------- */

  var timer = null
  var polling = false
  var getConfig = null
  var hooks = {}
  var state = new Map() // formID → { values, custom }
  var rendered = new Map() // formID → card element

  function container() {
    return document.getElementById('oc-form-cards')
  }

  function el(tag, className, text) {
    var node = document.createElement(tag)
    if (className) node.className = className
    if (text !== undefined) node.textContent = text
    return node
  }

  function shortSession(sessionID) {
    return String(sessionID || '').slice(0, 8)
  }

  function formState(form) {
    var s = state.get(form.id)
    if (!s) {
      s = { values: {}, custom: {} }
      for (const field of form.fields || []) {
        if (field.default === undefined) continue
        if (field.type === 'multiselect') s.values[field.key] = (field.default || []).slice()
        else s.values[field.key] = field.default
      }
      state.set(form.id, s)
    }
    return s
  }

  function refresh(form) {
    var card = rendered.get(form.id)
    var next
    if (card) {
      next = renderCard(form)
      card.replaceWith(next)
      // Keep the map pointing at the LIVE node: removeCard() must remove the
      // card the user actually sees, not a detached node left behind here.
      rendered.set(form.id, next)
    }
  }

  function renderCard(form) {
    var s = formState(form)
    var card = el('div', 'form-card')
    card.dataset.formId = form.id

    var head = el('div', 'form-card-head')
    head.appendChild(el('span', 'form-card-title', form.title || 'Question'))
    head.appendChild(el('span', 'form-card-session', shortSession(form.sessionID)))
    var dismiss = el('button', 'form-card-dismiss', '×')
    dismiss.type = 'button'
    dismiss.title = 'Dismiss (agent sees the question as dismissed)'
    dismiss.addEventListener('click', () => dismissForm(form))
    head.appendChild(dismiss)
    card.appendChild(head)

    var fields = visibleFields(form.fields || [], s.values)
    for (const field of fields) card.appendChild(renderField(form, field))

    var error = validate(form.fields || [], s)
    var foot = el('div', 'form-card-foot')
    var status = el('span', 'form-card-status', error || '')
    foot.appendChild(status)
    var submit = el('button', 'form-card-submit primary', 'Submit')
    submit.type = 'button'
    submit.disabled = !!error
    submit.addEventListener('click', () => submitForm(form))
    foot.appendChild(submit)
    card.appendChild(foot)
    return card
  }

  function renderField(form, field) {
    var s = formState(form)
    var wrap = el('div', 'form-field')
    var label = field.title || field.key
    if (field.description) wrap.title = field.description

    if (field.type === 'string' && (field.options || []).length) {
      wrap.appendChild(el('div', 'form-field-label', label))
      for (const option of field.options) {
        const id = `${form.id}-${field.key}-${option.value}`
        const row = el('label', 'form-option')
        const input = el('input', '')
        input.type = 'radio'
        input.name = `${form.id}-${field.key}`
        input.id = id
        input.checked = s.values[field.key] === option.value
        input.addEventListener('change', () => {
          s.values[field.key] = option.value
          refresh(form)
        })
        row.appendChild(input)
        const text = el('span', 'form-option-text')
        text.appendChild(el('span', 'form-option-label', option.label))
        if (option.description) text.appendChild(el('span', 'form-option-desc', option.description))
        row.appendChild(text)
        wrap.appendChild(row)
      }
      return wrap
    }

    if (field.type === 'multiselect') {
      wrap.appendChild(el('div', 'form-field-label', label))
      const picked = s.values[field.key] || []
      for (const option of field.options || []) {
        const mid = `${form.id}-${field.key}-${option.value}`
        const mrow = el('label', 'form-option')
        const minput = el('input', '')
        minput.type = 'checkbox'
        minput.id = mid
        minput.checked = picked.indexOf(option.value) !== -1
        minput.addEventListener('change', () => {
          var cur = (formState(form).values[field.key] || []).slice()
          if (minput.checked) {
            if (cur.indexOf(option.value) === -1) cur.push(option.value)
          } else {
            cur = cur.filter((v) => v !== option.value)
          }
          formState(form).values[field.key] = cur
          refresh(form)
        })
        mrow.appendChild(minput)
        const mtext = el('span', 'form-option-text')
        mtext.appendChild(el('span', 'form-option-label', option.label))
        if (option.description) mtext.appendChild(el('span', 'form-option-desc', option.description))
        mrow.appendChild(mtext)
        wrap.appendChild(mrow)
      }
      if (field.custom) {
        const other = el('input', 'form-input form-other')
        other.type = 'text'
        other.placeholder = 'Type your own answer…'
        other.value = s.custom[field.key] || ''
        other.addEventListener('input', () => {
          formState(form).custom[field.key] = other.value
          refreshKeepFocus(form, field.key)
        })
        wrap.appendChild(other)
      }
      return wrap
    }

    if (field.type === 'boolean') {
      wrap.appendChild(el('div', 'form-field-label', label))
      const brow = el('div', 'form-boolean')
      for (const choice of [
        { value: true, text: 'Yes' },
        { value: false, text: 'No' },
      ]) {
        const bbtn = el('button', `form-choice${s.values[field.key] === choice.value ? ' selected' : ''}`, choice.text)
        bbtn.type = 'button'
        bbtn.addEventListener('click', () => {
          formState(form).values[field.key] = choice.value
          refresh(form)
        })
        brow.appendChild(bbtn)
      }
      wrap.appendChild(brow)
      return wrap
    }

    if (field.type === 'external') {
      wrap.appendChild(el('div', 'form-field-label', label))
      const orow = el('div', 'form-boolean')
      const open = el('button', 'form-choice', 'Open link')
      open.type = 'button'
      open.addEventListener('click', () => window.open(field.url, '_blank'))
      orow.appendChild(open)
      const done = el('button', 'form-choice', 'Done')
      done.type = 'button'
      done.addEventListener('click', () => {
        formState(form).values[field.key] = true
        submitForm(form)
      })
      orow.appendChild(done)
      wrap.appendChild(orow)
      return wrap
    }

    // string (free text), number, integer
    wrap.appendChild(el('div', 'form-field-label', label))
    var input = el('input', 'form-input')
    input.type = field.type === 'number' || field.type === 'integer' ? 'number' : 'text'
    input.value = s.values[field.key] === undefined ? '' : String(s.values[field.key])
    if (field.placeholder) input.placeholder = field.placeholder
    input.addEventListener('input', () => {
      formState(form).values[field.key] = input.value
      refreshKeepFocus(form, field.key)
    })
    wrap.appendChild(input)
    return wrap
  }

  /* Text inputs re-render the card on every keystroke (validation state);
     restore focus and caret so typing stays continuous. */
  function refreshKeepFocus(form, key) {
    var active = document.activeElement
    var caret = active?.selectionStart
    var card = rendered.get(form.id)
    if (card) card.replaceWith(renderCard(form))
    if (active && active.tagName === 'INPUT') {
      const fresh = rendered.get(form.id)
      if (fresh) {
        const inputs = fresh.querySelectorAll('input[type="text"], input[type="number"]')
        for (const candidate of inputs) {
          if (candidate.value === active.value) {
            candidate.focus()
            if (caret !== undefined && candidate.setSelectionRange) candidate.setSelectionRange(caret, caret)
            break
          }
        }
      }
    }
    void key
  }

  function removeCard(formId) {
    state.delete(formId)
    var card = rendered.get(formId)
    if (card) card.remove()
    rendered.delete(formId)
  }

  function syncRendered() {
    // Drop cards for forms that settled server-side (answered/cancelled elsewhere).
    for (const id of Array.from(rendered.keys())) {
      if (!pendingIds.has(id)) removeCard(id)
    }
  }

  var pendingIds = new Set()
  /* Forms settled from this pane (reply accepted, or dismiss answered with
     OpenCode 409 FormAlreadySettled). render() never re-adds these. */
  var settledIds = new Set()

  function render(forms) {
    pendingIds = new Set((forms || []).map((f) => f.id))
    syncRendered()
    var box = container()
    if (!box) return
    for (const form of forms || []) {
      // Forms we settled ourselves (submitted or dismissed) can linger in the
      // server's form list for a while; re-rendering them would resurrect a
      // card the user just got rid of (observed live: OpenCode 409
      // FormAlreadySettled on the next dismiss).
      if (settledIds.has(form.id)) continue
      if (rendered.has(form.id)) continue
      const card = renderCard(form)
      rendered.set(form.id, card)
      box.appendChild(card)
    }
  }

  function submitForm(form) {
    var s = formState(form)
    var error = validate(form.fields || [], s)
    if (error) return
    var answer = buildAnswer(form.fields || [], s)
    OpenCodeChat.replyForm(getConfig(), form.sessionID, form.id, answer).then(
      () => {
        settledIds.add(form.id)
        removeCard(form.id)
        if (hooks.onReplied) hooks.onReplied(form.sessionID)
      },
      (err) => {
        // Already answered or cancelled elsewhere → the card is stale anyway.
        if (/409|FormAlreadySettled|FormNotFound/.test(String(err))) {
          settledIds.add(form.id)
          removeCard(form.id)
          return
        }
        const staleCard = rendered.get(form.id)
        if (staleCard) {
          const status = staleCard.querySelector('.form-card-status')
          if (status) status.textContent = String(err.message || err).slice(0, 200)
        }
      },
    )
  }

  function dismissForm(form) {
    var cfg
    try {
      cfg = getConfig() || OpenCodeChat.loadConfig()
    } catch (_e) {
      cfg = OpenCodeChat.loadConfig()
    }
    var p
    try {
      p = OpenCodeChat.cancelForm(cfg, form.sessionID, form.id, 'Dismissed from the Office taskpane')
    } catch (_e) {
      removeCard(form.id)
      return
    }
    p.then(
      () => {
        settledIds.add(form.id)
        removeCard(form.id)
      },
      (err) => {
        if (/409|FormAlreadySettled|FormNotFound/.test(String(err))) settledIds.add(form.id)
        removeCard(form.id)
      },
    )
  }

  /* Defensive polling: an exception inside getConfig/listForms/render used to
     leave the `polling` flag stuck at true, which froze the whole form list —
     cards then never disappeared (observed live). Every step now resets the
     flag and swallows its own errors. */

  function poll() {
    if (polling || !getConfig) return
    polling = true
    var cfg
    try {
      cfg = getConfig()
    } catch (_e) {
      polling = false
      return
    }
    var p
    try {
      p = OpenCodeChat.listForms(cfg)
    } catch (_e) {
      polling = false
      return
    }
    p.then(
      (forms) => {
        polling = false
        // A render error must never wedge the polling flag: the next tick
        // has to keep running or settled cards would linger forever.
        try {
          render(forms)
        } catch (_e) {
          /* skip this tick */
        }
      },
      () => {
        // OpenCode unreachable — keep existing cards, retry on the next tick.
        polling = false
      },
    )
  }

  function start(getConfigFn, hooksOpt) {
    getConfig = getConfigFn
    hooks = hooksOpt || {}
    if (timer) return
    poll()
    timer = setInterval(poll, POLL_MS)
  }

  function stop() {
    if (timer) clearInterval(timer)
    timer = null
  }

  return {
    visibleFields: visibleFields,
    buildAnswer: buildAnswer,
    validate: validate,
    start: start,
    stop: stop,
    render: render,
  }
})()

// Expose for the taskpane: chat.js reads the global. An explicit window
// assignment survives bundlers and strict module scopes too.
if (typeof window !== 'undefined') {
  window.OpenCodeForms = OpenCodeForms
}
