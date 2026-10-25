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

# install the add-in into the Office apps (needs Full Disk Access for the terminal)
node --experimental-strip-types scripts/sideload.ts                # PowerPoint
node --experimental-strip-types scripts/sideload.ts --app word     # Word
node --experimental-strip-types scripts/sideload.ts --app excel    # Excel
```

Then open a document in any of the three apps, click the add-in button in the Home ribbon, and point your MCP client at the server (below) — **no bridge to start manually**.

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

Verified on Windows 11 (ARM64, Office Insider 16.0.20430) in a Parallels VM. Recent Windows Office builds reject the old sideload shortcuts (Developer-registry manifests only load under a debugger, the new "Apps" storefront never queries custom catalogs), so the supported route is the classic add-in dialog with a **shared folder catalog**:

1. Copy `dist/index.cjs` and `addin/` to the Windows machine (e.g. `C:\o365-bridge`). You do **not** need to start anything: the first MCP client connect starts the daemon automatically (same stdio setup as above, with Windows paths). To run it standalone instead: `node C:\o365-bridge\dist\index.cjs --http --bridge` — taskpane on `http://localhost:8080`, MCP on `http://localhost:3001/mcp`.

2. Publish the `addin/` folder as a network share (UNC path) that the Windows user can list — a native SMB share, or on a VM the host's shared folder (e.g. `\\Mac\Home\office-live-mcp\addin` under Parallels).

3. Trust that folder as an add-in catalog (per user, one-time):

   ```bat
   reg add HKCU\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{7E3F2A11-93C4-4B5D-8E6F-1A2B3C4D5E6F} /v Id   /t REG_SZ   /d {7E3F2A11-93C4-4B5D-8E6F-1A2B3C4D5E6F} /f
   reg add HKCU\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{7E3F2A11-93C4-4B5D-8E6F-1A2B3C4D5E6F} /v Url  /t REG_SZ   /d \\host\share\addin /f
   reg add HKCU\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{7E3F2A11-93C4-4B5D-8E6F-1A2B3C4D5E6F} /v Flags /t REG_DWORD /d 1 /f
   ```

4. Restart the Office app, then: **Home → Add-ins → Advanced...** → **SHARED FOLDER** tab → select *Office Live MCP* → **Add**. The taskpane opens and connects to the bridge over `ws://localhost:8080`. The daemon's environment file lives at `%USERPROFILE%\.office-live-mcp\daemon.env`, its log at `%USERPROFILE%\.office-live-mcp\daemon.log`.

Notes:

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
