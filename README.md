<div align="center">

<img src="icon.png" width="96" alt="o365 MCP logo"/>

# o365 MCP

One MCP server that drives **live, open** Word, Excel and PowerPoint documents on macOS through Office.js.

</div>

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D24.0.0-brightgreen)](https://nodejs.org)

This started as a fork of [kzarzycki/powerpoint-mcp](https://github.com/kzarzycki/powerpoint-mcp) and grew into something bigger. The original idea is brilliant: instead of parsing files on disk, an Office.js add-in inside the running Office app talks to a local bridge server over WebSocket. Your AI assistant edits the document you have open right now, and you watch it happen.

I asked the original author whether he planned to extend this to Word and Excel ([issue #173](https://github.com/kzarzycki/powerpoint-mcp/issues/173)). While waiting for an answer I built it myself, on this branch. Same architecture, two more apps.

## What it does

One bridge server, one add-in taskpane, three Office apps. The add-in detects which host it runs in (PowerPoint, Word or Excel) and every tool call executes inside the live document.

- **PowerPoint**: 27 tools from the upstream project. Inspect decks, edit shapes and text, raw OOXML access, charts, screenshots, speaker notes, and an `execute_officejs` escape hatch.
- **Excel**: 24 new tools. Read and write ranges with values, formulas and number formats, inspect workbooks, format cells, manage worksheets, sort and find, tables, charts, conditional formatting, named ranges, data validation, comments, pivot refresh, sheet protection, plus `execute_exceljs`.
- **Word**: 19 new tools. Read and edit paragraphs, styles and formatting, search and replace, tables, comments, page breaks, tracked changes (read), selection, save, plus `execute_wordjs`.

That is 70 tools on one server. Anything the typed tools do not cover, the execute escape hatch handles with raw Office.js in the right host.

## Why live matters

File-based tools (openpyxl, python-docx, python-pptx) read and write bytes on disk. They cannot refresh a PivotTable, they silently drop charts and macros, and you only see the result after reopening the file. This server works inside the running app instead: formulas are calculated by Excel itself, edits appear on screen while the agent works, and nothing that Excel owns gets lost.

## Quick start

```bash
git clone -b o365 https://github.com/benjaminfrombe/powerpoint-mcp.git
cd powerpoint-mcp
npm install
npm run build

# install the add-in into the Office apps (needs Full Disk Access for the terminal)
node --experimental-strip-types scripts/sideload.ts                # PowerPoint
node --experimental-strip-types scripts/sideload.ts --app word     # Word
node --experimental-strip-types scripts/sideload.ts --app excel    # Excel

# start the bridge
node dist/index.cjs --http --bridge
```

Then open a document in any of the three apps, click the add-in button in the Home ribbon, and connect your MCP client to `http://127.0.0.1:3001/mcp`.

## How it is tested

Two layers, because they catch different bugs.

The mock suite (375 tests) verifies tool registration, argument handling and the generated Office.js code. It is fast and runs on every change.

The live suite runs every tool against a real open document. This is where the mock tests turned out to be insufficient: live testing caught three real API bugs the mocks could never find. `Range.findAllOrNullObject` does not exist on this host, `DataValidationRule` takes a type-keyed object rather than a `{ type, ... }` envelope, and calculation lives on `workbook.application`, not `workbook`. All fixed and now covered by both layers.

## Credits

This project stands on the work of others.

- [kzarzycki/powerpoint-mcp](https://github.com/kzarzycki/powerpoint-mcp) (MIT): the bridge architecture, the PowerPoint tool modules and the add-in are his. This fork exists because of it.
- [mindstone/mcp-servers](https://github.com/mindstone/mcp-servers) (FSL-1.1-MIT): studied as a reference for Word and Excel add-in design. No code from it is included, its license does not allow that, so the Word and Excel tools here are an independent implementation.
- The live Word MCP projects ([kennypeh85/word-mcp](https://github.com/kennypeh85/word-mcp), [HelloWorld-Open/word-mcp-server](https://github.com/HelloWorld-Open/word-mcp-server), [Han-ops-dev/word-mcp-server](https://github.com/Han-ops-dev/word-mcp-server), all MIT) and [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) (MIT) shaped which tools were worth building. Design references only, no code.
- The icon uses the [Tabler Icons](https://github.com/tabler/tabler-icons) bridge glyph as a starting point for the ring mark (MIT).

Full attribution in [NOTICE](NOTICE).

## Status

Working on macOS with desktop Word, Excel and PowerPoint. PowerPoint Web works for the upstream tool set. The Word and Excel tool surfaces are young: they cover what I actually needed, and the escape hatch covers the rest. Expect rough edges and please open issues on the fork.

## License

MIT, like the upstream project. See [LICENSE](LICENSE).
