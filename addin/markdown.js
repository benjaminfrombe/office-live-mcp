/* Markdown for assistant bubbles, rendered by markdown-it (vendored in
   addin/vendor, the same parser Paseo uses): CommonMark plus GFM tables and
   strikethrough.

   Safety: html is off, so raw HTML in model output is escaped as text, and
   markdown-it's validateLink already refuses javascript:/vbscript:/data:
   URLs. The HTML string it returns is therefore safe to assign. */
var OfficeMarkdown = (() => {
  var md = window.markdownit({
    html: false,
    linkify: true, // bare URLs become links, as in every chat UI
    breaks: true, // a single newline is a line break: models write that way
    typographer: false,
  })

  // Links open outside the taskpane, never inside it.
  var defaultLinkOpen =
    md.renderer.rules.link_open || ((tokens, idx, options, _env, self) => self.renderToken(tokens, idx, options))
  md.renderer.rules.link_open = (tokens, idx, options, env, self) => {
    tokens[idx].attrSet('target', '_blank')
    tokens[idx].attrSet('rel', 'noopener noreferrer')
    return defaultLinkOpen(tokens, idx, options, env, self)
  }

  // Tables get a wrapper that scrolls sideways: the pane is ~320px wide.
  md.renderer.rules.table_open = () => '<div class="md-table-wrap"><table class="md-table">\n'
  md.renderer.rules.table_close = () => '</table></div>\n'

  /* Markdown text to a DocumentFragment. */
  function render(text) {
    var template = document.createElement('template')
    template.innerHTML = md.render(String(text || ''))
    return template.content
  }

  return { render: render }
})()

// Expose for the taskpane: chat.js reads the global (same idiom as opencode.js).
if (typeof window !== 'undefined') {
  window.OfficeMarkdown = OfficeMarkdown
}
