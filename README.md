<div align="center">

<img src="icon.png" width="96" alt="Office Live MCP logo"/>

# Office Live MCP

One MCP server that drives **live, open** Word, Excel and PowerPoint documents on macOS and Windows through Office.js.

</div>

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D24.0.0-brightgreen)](https://nodejs.org)

This started as a fork of [kzarzycki/powerpoint-mcp](https://github.com/kzarzycki/powerpoint-mcp) and grew into something bigger. The original idea is brilliant: instead of parsing files on disk, an Office.js add-in inside the running Office app talks to a local bridge server over WebSocket. Your AI assistant edits the document you have open right now, and you watch it happen.

I asked the original author whether he planned to extend this to Word and Excel ([issue #173](https://github.com/kzarzycki/powerpoint-mcp/issues/173)). While waiting for an answer I built it myself, on this branch. Same architecture, two more apps.

## What it does

One bridge server, one add-in taskpane, three Office apps. The add-in detects which host it runs in (PowerPoint, Word or Excel) and every tool call executes inside the live document.

- **PowerPoint**: 27 tools from the upstream project. Inspect decks, edit shapes and text, raw OOXML access, charts, screenshots, speaker notes, and an `execute_officejs` escape hatch.
- **Excel**: 32 new tools. Read and write ranges with values, formulas and number formats, inspect workbooks, format cells, manage worksheets, sort and find, tables, charts, conditional formatting, named ranges, data validation, comments, pivot refresh, sheet protection, plus `execute_exceljs`.
- **Word**: 26 new tools. Read and edit paragraphs, styles and formatting, search and replace, tables, comments, page breaks, tracked changes (read), selection, save, plus `execute_wordjs`.

That is 86 tools on one server. Anything the typed tools do not cover, the execute escape hatch handles with raw Office.js in the right host.

## Why live matters

File-based tools (openpyxl, python-docx, python-pptx) read and write bytes on disk. They cannot refresh a PivotTable, they silently drop charts and macros, and you only see the result after reopening the file. This server works inside the running app instead: formulas are calculated by Excel itself, edits appear on screen while the agent works, and nothing that Excel owns gets lost.

There is a second advantage that matters just as much. Everything here runs locally on your machine, next to your agent, not inside some hosted sandbox. Cloud Office agents lock you into their environment: your files go up, the tools they offer are the tools you get, and combining them with anything else is a fight against the platform. Here the Office tools are just one part of what the agent can reach. In the same session it can read your mail, query a database, drive a browser, call any other MCP server or plain script, and then put the results straight into the live document. The Office control is a capability, not a cage. That combination is where most of the real value showed up while building and testing this.

## Quick start

```bash
git clone https://github.com/benjaminfrombe/office-live-mcp.git
cd office-live-mcp
npm install
npm run build

# install the add-in into the Office apps + print your MCP client config
node dist/index.cjs setup
```

`setup` is idempotent — run it again any time. On macOS it sideloads the manifest into each installed Office app (needs Full Disk Access for your terminal; it tells you if macOS blocks it). Then open a document, click the add-in button in the Home ribbon, and connect your MCP client with the printed config. **No bridge to start manually.**

> With an npm install (`npm i -g office-live-mcp`) the same command is just `office-live-mcp setup`.

## Connecting an MCP client

The default mode is **stdio**: the client (Claude Code, OpenCode, Claude Desktop, …) spawns the server, and the server makes sure a shared bridge daemon is running.

```json
{
  "mcpServers": {
    "office-live-mcp": { "command": "node", "args": ["/absolute/path/to/office-live-mcp/dist/index.cjs"] }
  }
}
```

What happens on connect:

- If no daemon is up, one is started **detached** (`--http --bridge`: taskpane on :8080, MCP on :3001) and it stays up after the client session ends — the Office add-in needs it even when no AI session is open.
- If a daemon is already up (another harness, an earlier session, the machine just rebooted), it is reused. Multiple harnesses share one daemon and see the same live documents.
- If the daemon crashed or died with a reboot, the next client connect brings it back.

Useful extras:

- `node dist/index.cjs --stop-daemon` — stop the shared daemon (however it was started).
- `~/.office-live-mcp/daemon.log` — daemon log; `~/.office-live-mcp/daemon.env` — optional `KEY=VALUE` lines merged into the daemon environment (e.g. `BRIDGE_TLS=1`).
- Prefer an HTTP MCP client? Start the daemon once (`node dist/index.cjs --http --bridge`) and connect to `http://127.0.0.1:3001/mcp`.
- Legacy single-session mode (`--stdio --bridge`, everything in one process) still exists but does not survive multiple concurrent clients.

### Optional: the ribbon icon (HTTPS)

Everything works over plain HTTP except one cosmetic detail: Office refuses to load the **ribbon icon** over HTTP, so the button shows a generic placeholder. The add-in, taskpane and bridge connection are unaffected.

To get the real icon, serve the bridge over HTTPS with a locally trusted certificate (one-time setup):

```bash
brew install mkcert && mkcert -install        # or: npx office-addin-dev-certs install
mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost 127.0.0.1 ::1

printf 'BRIDGE_TLS=1\n' > ~/.office-live-mcp/daemon.env   # picked up by the auto-started daemon
node --experimental-strip-types scripts/sideload.ts --tls      # per app: add --app word / --app excel
```

Then restart the Office apps. Without this step, just ignore the placeholder icon.

## Windows installation

Verified on Windows 11 (ARM64, Office Insider 16.0.20430) in a Parallels VM. Recent Windows Office builds reject the old sideload shortcuts (Developer-registry manifests only load under a debugger, the new "Apps" storefront never queries custom catalogs, and neither the classic nor the new dialog offers an "Upload My Add-in" option), so the supported route is a **shared folder catalog**. `setup` automates all of it:

```bat
node C:\path\to\office-live-mcp\dist\index.cjs setup
```

What it does:

1. Copies `dist/` and `addin/` to `%LOCALAPPDATA%\office-live-mcp`, keeping only the manifest variants that match the daemon transport (HTTPS or plain) so the catalog lists exactly one add-in per Office app.
2. Registers a trusted add-in catalog in the registry (HKCU, no elevation needed) pointing at that folder through a UNC path. Office only trusts network paths for custom add-ins — that is a Microsoft requirement. In practice no share has to be created: setup uses, in order, an `OLM_CATALOG_UNC` override → the Parallels host share when running from `\\Mac\Home\...` → a dedicated `OfficeLiveMCP` share (created automatically in an elevated terminal) → the built-in admin share (`\\localhost\C$`), which exists by default on admin machines.
3. Prints the ready-to-paste MCP client config.

Then the one step no installer can do, because Office requires it:

4. **Close and reopen the Office app** → **Home → Add-ins → Advanced...** → **SHARED FOLDER** tab → select *Office Live MCP* → **Add**.

> **Why not an HTTPS catalog (no UNC at all)?** The bridge can emulate a SharePoint app catalog over `https://localhost:8443` (`OLM_HTTPS_CATALOG=1`, see `server/sharepoint-catalog.ts`), and Office would then list the add-in under MY ORGANIZATION. This worked on Office builds up to 16.0.20430, but Microsoft fully retired SharePoint add-in catalogs on April 2, 2026 — current builds still probe the catalog yet discard the results, so the route is dead on modern Office. The UNC route above is the only supported one.

Notes:

- The daemon's environment file lives at `%USERPROFILE%\.office-live-mcp\daemon.env`, its log at `%USERPROFILE%\.office-live-mcp\daemon.log`.
- Use `manifest.xml` for PowerPoint, `manifest.word.xml` for Word, `manifest.excel.xml` for Excel. All point the taskpane at `http://localhost:8080`.
- As on macOS, the ribbon icon stays a placeholder unless the bridge runs over HTTPS (see "Optional: the ribbon icon" above) and the manifests point at `https://localhost:8443`.
- If a build refuses plain-HTTP catalogs, the bridge can also emulate a SharePoint-style HTTPS catalog (`server/sharepoint-catalog.ts`, enabled with `BRIDGE_TLS=1`): Office enumerates `https://<host>:8443` as a corporate catalog and the same manifests are served from there.

## How it is tested

Two layers, because they catch different bugs.

The mock suite (411 tests) verifies tool registration, argument handling and the generated Office.js code, plus the stdio→daemon proxy (auto-start, session handling, SSE parsing, stop). It is fast and runs on every change.

The live suite runs every tool against a real open document. This is where the mock tests turned out to be insufficient: live testing caught three real API bugs the mocks could never find. `Range.findAllOrNullObject` does not exist on this host, `DataValidationRule` takes a type-keyed object rather than a `{ type, ... }` envelope, and calculation lives on `workbook.application`, not `workbook`. All fixed and now covered by both layers.

## Credits

This project stands on the work of others.

- [kzarzycki/powerpoint-mcp](https://github.com/kzarzycki/powerpoint-mcp) (MIT): the bridge architecture, the PowerPoint tool modules and the add-in are his. This fork exists because of it.
- [mindstone/mcp-servers](https://github.com/mindstone/mcp-servers) (FSL-1.1-MIT): studied as a reference for Word and Excel add-in design. No code from it is included, its license does not allow that, so the Word and Excel tools here are an independent implementation.
- The live Word MCP projects ([kennypeh85/word-mcp](https://github.com/kennypeh85/word-mcp), [HelloWorld-Open/word-mcp-server](https://github.com/HelloWorld-Open/word-mcp-server), [Han-ops-dev/word-mcp-server](https://github.com/Han-ops-dev/word-mcp-server), all MIT) and [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) (MIT) shaped which tools were worth building. Design references only, no code.
- The icon uses the [Tabler Icons](https://github.com/tabler/tabler-icons) bridge glyph as a starting point for the ring mark (MIT).

Full attribution in [NOTICE](NOTICE).

## Status

Working on macOS with desktop Word, Excel and PowerPoint, and verified on Windows 11 ARM64 (Parallels) with PowerPoint: stdio auto-daemon, OpenCode as MCP client, and live edits through the add-in. PowerPoint Web works for the upstream tool set. The Word and Excel tool surfaces are young: they cover what I actually needed, and the escape hatch covers the rest. Expect rough edges and please open issues.

## License

MIT, like the upstream project. See [LICENSE](LICENSE).
