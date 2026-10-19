import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'

/**
 * Excel tools — scaffold for the o365 extension.
 *
 * The bridge, connection pool and executeCode escape hatch are host-neutral:
 * an Excel add-in connecting to the same bridge registers as a connection
 * whose code runs inside Excel.run(). Deep tools (ranges, formulas, charts,
 * tables) will be added incrementally following the patterns in
 * inspect.ts/charts.ts/text.ts.
 */
export function registerExcelTools(
  server: McpServer,
  pool: ConnectionPool,
  _getSessionId: () => string | undefined,
  _getActiveSessionCount: () => number,
): void {
  server.tool(
    'list_workbooks',
    'Lists all Excel workbooks currently connected to the bridge server. Shows workbook IDs (file paths for saved files, generated IDs for unsaved) and connection status. Use this to find the workbookId to pass to other Excel tools when multiple workbooks are open.',
    withTool(async () => {
      const workbooks = []
      for (const [id, conn] of pool.entries()) {
        workbooks.push({
          workbookId: id,
          filePath: conn.filePath,
          ready: conn.ready,
        })
      }
      return {
        content: [
          {
            type: 'text' as const,
            text:
              workbooks.length === 0
                ? 'No workbooks connected. Open a workbook in Excel and click the add-in button.'
                : JSON.stringify(workbooks, null, 2),
          },
        ],
      }
    }),
  )

  server.tool(
    'execute_exceljs',
    "Execute arbitrary Office.js code inside the live Excel workbook. The code runs inside Excel.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. Common operations: read/write ranges (workbook.worksheets), formulas, tables, charts. Cell references use A1 notation; values/formulas are returned as arrays matching the range shape.",
    {
      code: z
        .string()
        .describe(
          "Office.js code to execute. Runs inside Excel.run() with 'context' available. Use 'return' to send back a result.",
        ),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async ({ code, workbookId }) => {
      const target = pool.resolveTarget(workbookId)
      const result = await pool.sendCommand('executeCode', { code }, target.ws)
      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result ?? { success: true }, null, 2) }],
      }
    }),
  )
}
