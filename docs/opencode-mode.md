# OpenCode Chat Mode (reverse bridge)

The taskpane has two modes, switchable with the tabs at the top:

- **Bridge** — the original direction: the MCP server drives Office via WS commands (`app.js`).
- **OpenCode Chat** — the reverse direction: the taskpane connects to the OpenCode HTTP
  server as a client, renders a chat, and applies replies to the live document via
  Office.js (`opencode.js` + `chat.js`). The bridge keeps running in the background
  while the chat tab is open.

## Setup

The OpenCode API always requires basic auth with a per-service password (there
is no flag to disable it), and the service picks a random port and password on
every start. To keep the add-in passwordless and stable, run the local proxy
(`scripts/opencode-proxy.ts`):

```sh
npm run opencode-proxy
# or: node --experimental-strip-types ./scripts/opencode-proxy.ts
```

It listens on `http://127.0.0.1:4096` (override with `OPENCODE_PROXY_PORT`),
reads the real URL + password from `~/.local/state/opencode/service.json`,
injects the `Authorization` header, and re-reads the file every 2s so service
restarts are picked up automatically. Streaming (SSE) passes through chunk by
chunk.

Then in the taskpane: open the **OpenCode Chat** tab. The password field is
hidden by default (the proxy handles auth). **In Office, set the URL to
`https://127.0.0.1:4096`** — Office's WKWebView blocks mixed content, so an
HTTPS-served taskpane cannot fetch plain `http://` URLs (that is what a
"Load failed" status means here). In a plain browser test `http://` works too.

Manual (without proxy) configuration is still possible: `cat
~/.local/state/opencode/service.json` gives the direct URL and password, both
stored in `localStorage` (`opencode-chat-config`).

Buttons:

- **+ Selection** — pastes the current document selection into the input as context.
- **Insert last reply at selection** — writes the last assistant reply into the
  document at the selection (`setSelectedDataAsync`, plain text).

## Image attachments (paste)

Pasting an image into the chat input adds it as a pending attachment, shown as
a thumbnail chip (with a ✕ remove button) above the input. On send the images
travel inline in the prompt request as data URIs — `POST /api/session/{id}/prompt`
gains a `files: [{uri, source: {type: "inline"}, name}]` array
(`promptFiles` in `opencode.js`). Images over 8 MB of base64 are rejected with
a status message; attachments are cleared only after a successful send so a
failed send can be retried. Note: if WKWebView does not deliver image paste
items, pasting silently does nothing (no fallback).

## Reasoning strip

A thin collapsible strip sits at the very top of the chat area (below the
session toolbar). While the assistant works, it streams the live reasoning
tokens (`session.reasoning.started/delta/ended` are added to the SSE
subscription; deltas buffer into `chatState.reasoningText`, reset on
`session.step.started`) auto-scrolled to the bottom, monospace and dimmed.
When the step ends, the strip collapses to one line —
"Reasoning · N chars · click to expand" — and clicking toggles the full text.
The strip is hidden whenever there is no reasoning content.

## How it talks to OpenCode

All requests are plain `fetch` with HTTP basic auth (`opencode:<password>`),
no build step and no `@opencode/client` dependency (the taskpane is unbundled
static files):

| Action | Call |
|---|---|
| List sessions | `GET /api/session` |
| New session | `POST /api/session` `{title}` |
| History | `GET /api/session/{id}/message` |
| Send prompt | `POST /api/session/{id}/prompt` `{text}` |
| Live updates | `GET /api/event?types=...` (SSE) |

SSE details:

- `EventSource` cannot send basic-auth headers, so the stream is consumed through
  `fetch` + `ReadableStream` (`subscribeEvents` in `opencode.js`).
- Subscribed types: `session.step.*`, `session.text.delta`, `session.text.ended`,
  `session.reasoning.*`, `session.execution.*`. `session.text.delta` appends to the
  streaming bubble; `session.reasoning.delta` feeds the reasoning strip; terminal
  events trigger a full history refetch so the pane always matches the server.
- Streams are live-only (no replay): on close/error the client resubscribes after 2s.
- Responses are wrapped in a `data` envelope; assistant messages carry a `content`
  array of typed parts, user messages a plain `text` field.

## Testing

`addin/opencode.test.ts` runs the client in a `node:vm` sandbox (same harness
pattern as `app.test.ts`), covering config storage, auth headers, the data
envelope, message-text extraction, and SSE chunk parsing. Run with `npm run check`.

## Not yet implemented (phase 2)

- Tool loop: agent tool calls executed by the taskpane via Office.js (the full
  reverse of the MCP bridge). Would need an OpenCode plugin exposing RPC tools
  that proxy back to the taskpane.
- Streaming into the document (apply deltas as they arrive instead of on completion).
- Markdown rendering of replies.
