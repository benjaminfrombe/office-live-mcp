import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'

/**
 * Word tools — scaffold for the o365 extension.
 *
 * The bridge, connection pool and executeCode escape hatch are host-neutral:
 * a Word add-in connecting to the same bridge registers as a connection whose
 * code runs inside Word.run(). These tools give Word-specific descriptions;
 * deep inspection/editing tools (paragraphs, tables, tracked changes) will be
 * added incrementally following the patterns in inspect.ts/text.ts/xml.ts.
 */
export function registerWordTools(
  server: McpServer,
  pool: ConnectionPool,
  _getSessionId: () => string | undefined,
  _getActiveSessionCount: () => number,
): void {
  server.tool(
    'list_documents',
    'Lists all Word documents currently connected to the bridge server. Shows document IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the documentId to pass to other Word tools when multiple documents are open.',
    withTool(async () => {
      const documents = []
      for (const [id, conn] of pool.entries()) {
        documents.push({
          documentId: id,
          filePath: conn.filePath,
          ready: conn.ready,
        })
      }
      return {
        content: [
          {
            type: 'text' as const,
            text:
              documents.length === 0
                ? 'No documents connected. Open a document in Word and click the add-in button.'
                : JSON.stringify(documents, null, 2),
          },
        ],
      }
    }),
  )

  server.tool(
    'execute_wordjs',
    "Execute arbitrary Office.js code inside the live Word document. The code runs inside Word.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. Common operations: read the body (context.document.body), insert paragraphs, search and replace, work with tables, read/set tracked changes. All units are points.",
    {
      code: z
        .string()
        .describe(
          "Office.js code to execute. Runs inside Word.run() with 'context' available. Use 'return' to send back a result.",
        ),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async ({ code, documentId }) => {
      const target = pool.resolveTarget(documentId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)
      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result ?? { success: true }, null, 2) }],
      }
    }),
  )
}
