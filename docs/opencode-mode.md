# OpenCode Chat Mode (reverse bridge)

The taskpane has two modes, switchable with the tabs at the top:

- **Bridge** — the original direction: the MCP server drives Office via WS commands (`app.js`).
- **OpenCode Chat** — the reverse direction: the taskpane connects to the OpenCode HTTP
  server as a client, renders a chat, and applies replies to the live document via
  Office.js (`opencode.js` + `chat.js`). The bridge keeps running in the background
  while the chat tab is open.

## Setup

1. The OpenCode background service listens on a random port with basic auth. Get both:
   ```sh
   cat ~/.local/state/opencode/service.json
   # {"url": "http://127.0.0.1:49374", "password": "...", ...}
   ```
2. In the taskpane, open the **OpenCode Chat** tab, fill in the URL and password,
   press **Save & connect**. Both are stored in `localStorage`
   (`opencode-chat-config`).
3. Pick a session or press **New session**, then chat.

Buttons:

- **+ Selection** — pastes the current document selection into the input as context.
- **Insert last reply at selection** — writes the last assistant reply into the
  document at the selection (`setSelectedDataAsync`, plain text).

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
  `session.execution.*`. `session.text.delta` appends to the streaming bubble;
  terminal events trigger a full history refetch so the pane always matches the server.
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
