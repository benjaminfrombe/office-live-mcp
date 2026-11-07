import { execFile } from 'node:child_process'
import { existsSync } from 'node:fs'
import { extname } from 'node:path'
import { platform } from 'node:process'
import { promisify } from 'node:util'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import { withTool } from '../tool-helpers.ts'

const execFileAsync = promisify(execFile)

// ---------------------------------------------------------------------------
// open_document — launch a document file in its Office app.
//
// The live tools only reach documents that are ALREADY open in the Office
// app with the taskpane connected. Building a complete new deck live,
// slide by slide, is slow and limited. The intended flow for "make me a
// full presentation" is: generate the entire .pptx offline (python-pptx,
// OOXML, ...), then open_document it — the taskpane auto-opens in the new
// document and connects to the bridge, and live refinement can continue.
// ---------------------------------------------------------------------------

const DOCUMENT_EXTENSIONS = new Set(['.pptx', '.ppt', '.docx', '.doc', '.xlsx', '.xls'])

/** Pure validation shared with the unit tests. */
export function validateDocumentPath(inputPath: string): { ext: string } {
  if (!inputPath || typeof inputPath !== 'string') throw new Error('A file path is required.')
  const ext = extname(inputPath).toLowerCase()
  if (!DOCUMENT_EXTENSIONS.has(ext)) {
    throw new Error(
      `Unsupported file type "${ext}". Expected a PowerPoint, Word or Excel document (.pptx/.ppt/.docx/.doc/.xlsx/.xls).`,
    )
  }
  if (!existsSync(inputPath)) {
    throw new Error(`File not found: ${inputPath}. Generate the document first, then open it.`)
  }
  return { ext }
}

/** Platform launcher, injectable for tests. */
export type Launcher = (filePath: string) => Promise<void>

export const defaultLauncher: Launcher = async (filePath: string) => {
  if (platform === 'darwin') {
    // `open` picks the default app for the extension (PowerPoint/Word/Excel).
    await execFileAsync('open', [filePath])
    return
  }
  if (platform === 'win32') {
    // start "" "<path>" — the empty title argument is required because the
    // first quoted argument would otherwise be treated as the window title.
    await execFileAsync('cmd', ['/c', 'start', '', filePath])
    return
  }
  throw new Error(`Opening documents is not supported on platform "${platform}".`)
}

export function registerDocumentTools(
  server: McpServer,
  _pool: unknown,
  _getSessionId: () => string | undefined,
  _getActiveSessionCount: () => number,
  launcher: Launcher = defaultLauncher,
): void {
  server.tool(
    'open_document',
    'Open a document file (.pptx/.docx/.xlsx) in its Office app (PowerPoint, Word or Excel) on the machine where the bridge runs. ' +
      'Use this after generating a complete document offline: the taskpane auto-opens in the new document and connects to the bridge, ' +
      'after which the live tools can inspect and refine it. This is the intended route for "create a full presentation/document" ' +
      'requests — do not build complete decks live, slide by slide.',
    {
      path: z.string().describe('Absolute path to the document file. It must already exist on this machine.'),
    },
    withTool(async ({ path: filePath }: { path: string }) => {
      validateDocumentPath(filePath)
      await launcher(filePath)
      return {
        content: [
          {
            type: 'text' as const,
            text:
              `Opened ${filePath} in its Office app. The taskpane auto-opens in the new document and connects to the bridge ` +
              `(this can take a few seconds). Then use list_presentations / list_documents / list_workbooks to confirm the ` +
              `connection before driving it with the live tools.`,
          },
        ],
      }
    }),
  )
}
