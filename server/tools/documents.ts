import { execFile } from 'node:child_process'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { extname } from 'node:path'
import { platform } from 'node:process'
import { promisify } from 'node:util'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { errorMessage, withTool } from '../tool-helpers.ts'
import { buildInsertOptions, getConcurrentWarning } from './shared.ts'

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
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
  launcher: Launcher = defaultLauncher,
): void {
  server.tool(
    'open_document',
    'Open a document file (.pptx/.docx/.xlsx) as a NEW document in its Office app (PowerPoint, Word or Excel) on the machine where the bridge runs. ' +
      'Only use this when the user explicitly wants a separate new document/window — it does NOT fill a presentation they already have open. ' +
      "To deliver a built deck into the user's open presentation, use import_presentation_file instead.",
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

  server.tool(
    'import_presentation_file',
    'Import ALL slides from a .pptx file on disk into a connected live PowerPoint presentation, in one atomic operation. ' +
      'The file is read server-side (its Base64 never enters the model context) and inserted via insertSlidesFromBase64 — ' +
      'the slides appear live in the open presentation while the user watches. ' +
      'This is the route for "build me a full presentation INTO this open deck": generate the complete .pptx offline ' +
      '(e.g. python-pptx), then import it into the connected presentation. No second pane or ribbon click needed.',
    {
      path: z.string().describe('Absolute path to the source .pptx file. It must already exist on this machine.'),
      presentationId: z
        .string()
        .optional()
        .describe(
          'Destination presentation ID from list_presentations. Optional when only one presentation is connected.',
        ),
      formatting: z
        .enum(['KeepSourceFormatting', 'UseDestinationTheme'])
        .optional()
        .describe('Formatting mode for the inserted slides. Default: KeepSourceFormatting.'),
    },
    withTool(
      async ({
        path: filePath,
        presentationId,
        formatting,
      }: {
        path: string
        presentationId?: string
        formatting?: 'KeepSourceFormatting' | 'UseDestinationTheme'
      }) => {
        if (!filePath || typeof filePath !== 'string') throw new Error('A file path is required.')
        if (extname(filePath).toLowerCase() !== '.pptx') {
          throw new Error(`Unsupported file type "${extname(filePath)}". Only .pptx can be imported.`)
        }
        if (!existsSync(filePath)) {
          throw new Error(`File not found: ${filePath}. Generate the presentation first, then import it.`)
        }
        const size = statSync(filePath).size
        if (size > 200 * 1024 * 1024) {
          throw new Error(`File too large to import: ${Math.round(size / 1024 / 1024)} MB (max 200 MB).`)
        }
        const base64 = readFileSync(filePath).toString('base64')

        const dest = pool.resolveTarget(presentationId)
        const optionsArg = buildInsertOptions(formatting)
        const insertCode = `
          context.presentation.insertSlidesFromBase64("${base64}"${optionsArg});
          await context.sync();
          var slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return { slideCount: slides.items.length };
        `
        const inserted = (await pool.sendCommand('executeCode', { code: insertCode }, dest.ws)) as {
          slideCount: number
        }

        const warning = getConcurrentWarning(getSessionId(), dest.presentationId, getActiveSessionCount())
        const text =
          `Imported ${filePath} into the connected presentation — it now has ${inserted.slideCount} slides (the user saw them appear live).` +
          (warning ?? '')
        return { content: [{ type: 'text' as const, text }] }
      },
    ),
  )
}
