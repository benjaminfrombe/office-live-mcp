import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'
import { ConnectionPool } from '../bridge.ts'
import { registerTools } from '../tools.ts'

vi.mock('node:fs', () => ({ existsSync: vi.fn(() => true), readFileSync: vi.fn(), writeFileSync: vi.fn() }))

function mockWs(): WebSocket {
  return { send: vi.fn(), readyState: 1 } as unknown as WebSocket
}

async function waitForSend(ws: WebSocket, index: number) {
  const sendMock = ws.send as unknown as ReturnType<typeof vi.fn>
  await vi.waitFor(() => {
    if (sendMock.mock.calls.length <= index) {
      throw new Error(`ws.send call #${index} not sent yet`)
    }
  })
  return JSON.parse(sendMock.mock.calls[index][0])
}

async function setupMcpClient(pool: ConnectionPool) {
  const server = new McpServer({ name: 'test', version: '0.0.1' })
  registerTools(
    server,
    pool,
    () => 'test-session',
    () => 1,
  )

  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  await server.connect(serverTransport)

  const client = new Client({ name: 'test-client', version: '0.0.1' })
  await client.connect(clientTransport)

  return { client, server }
}

type CallResult = { text: string; isError?: boolean }

describe('Excel tools', () => {
  let pool: ConnectionPool
  let ws: WebSocket

  beforeEach(() => {
    pool = new ConnectionPool(100)
    ws = mockWs()
    pool.add('book.xlsx', {
      ws,
      ready: true,
      presentationId: 'book.xlsx',
      filePath: '/path/book.xlsx',
    })
  })

  /** Call a tool, capture the generated Office.js code, and answer it with `responseData`. */
  async function callTool(
    name: string,
    args: Record<string, unknown>,
    responseData?: unknown,
  ): Promise<{
    code: string
    result: CallResult
  }> {
    const { client } = await setupMcpClient(pool)
    const callPromise = client.callTool({ name, arguments: args })
    const sent = await waitForSend(ws, 0)
    expect(sent.action).toBe('executeCode')
    if (responseData !== undefined) {
      pool.handleResponse(sent.id, 'response', responseData)
    }
    const result = (await callPromise) as { content: Array<{ text: string }>; isError?: boolean }
    return {
      code: sent.params.code as string,
      result: { text: result.content[0].text, isError: result.isError },
    }
  }

  /** Call a tool that must fail locally (no ws round-trip). */
  async function callToolError(name: string, args: Record<string, unknown>): Promise<CallResult> {
    const { client } = await setupMcpClient(pool)
    const result = (await client.callTool({ name, arguments: args })) as {
      content: Array<{ text: string }>
      isError?: boolean
    }
    return { text: result.content[0].text, isError: result.isError }
  }

  // -------------------------------------------------------------------------

  describe('list_workbooks', () => {
    it('returns empty message with no connections', async () => {
      const emptyPool = new ConnectionPool(100)
      const { client } = await setupMcpClient(emptyPool)
      const result = (await client.callTool({ name: 'list_workbooks', arguments: {} })) as {
        content: Array<{ text: string }>
      }
      expect(result.content[0].text).toContain('No workbooks connected')
    })

    it('lists connected workbooks', async () => {
      const { client } = await setupMcpClient(pool)
      const result = (await client.callTool({ name: 'list_workbooks', arguments: {} })) as {
        content: Array<{ text: string }>
      }
      const parsed = JSON.parse(result.content[0].text)
      expect(parsed).toHaveLength(1)
      expect(parsed[0]).toEqual({ workbookId: 'book.xlsx', filePath: '/path/book.xlsx', ready: true })
    })
  })

  describe('execute_exceljs', () => {
    it('passes raw code through and returns the result', async () => {
      const { code, result } = await callTool(
        'execute_exceljs',
        { code: 'return context.workbook.name', workbookId: 'book.xlsx' },
        'MyBook',
      )
      expect(code).toContain('return context.workbook.name')
      expect(JSON.parse(result.text)).toBe('MyBook')
    })

    it('rejects an unknown workbookId', async () => {
      const r = await callToolError('execute_exceljs', { code: 'return 1', workbookId: 'nope.xlsx' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('not found')
    })
  })

  describe('read_range', () => {
    it('generates a load of values, formulas and dimensions', async () => {
      const { code, result } = await callTool(
        'read_range',
        { range: 'A1:C3', workbookId: 'book.xlsx' },
        { address: 'Sheet1!A1:C3', rowCount: 3, columnCount: 3, values: [[1, 2, 3]] },
      )
      expect(code).toContain('sheet.getRange(args.range)')
      expect(code).toContain('rowCount,columnCount,values')
      expect(code).toContain('"formulas"')
      expect(JSON.parse(result.text).values).toEqual([[1, 2, 3]])
    })

    it('injects sheetName as JSON, not interpolated', async () => {
      const { code } = await callTool('read_range', { range: 'A1', sheetName: 'Q3 Data' }, {})
      expect(code).toContain('var args = {"range":"A1","sheetName":"Q3 Data"}')
      expect(code).toContain('wbSheets.getItem(args.sheetName)')
    })

    it('uses the used range when requested', async () => {
      const { code } = await callTool('read_range', { usedRange: true }, {})
      expect(code).toContain('sheet.getUsedRange()')
    })
  })

  describe('write_range', () => {
    it('writes values and formulas and returns the written address', async () => {
      const { code, result } = await callTool(
        'write_range',
        {
          range: 'A1:B2',
          values: [
            [1, 2],
            [3, 4],
          ],
          formulas: [
            ['=A1', '=B1'],
            ['=A2', '=B2'],
          ],
          workbookId: 'book.xlsx',
        },
        { written: 'Sheet1!A1:B2', rowCount: 2, columnCount: 2 },
      )
      expect(code).toContain('range.values = args.values')
      expect(code).toContain('range.formulas = args.formulas')
      expect(JSON.parse(result.text).written).toBe('Sheet1!A1:B2')
    })

    it('rejects a write with no payload', async () => {
      const r = await callToolError('write_range', { range: 'A1' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('at least one of')
    })

    it('keeps hostile cell content inside the JSON literal', async () => {
      const hostile = '"); while(true) {} var x = ("'
      const { code } = await callTool('write_range', { range: 'A1', values: hostile }, {})
      // The value must appear JSON-escaped inside the args literal, so it can
      // never terminate the string and inject code into the Office.js body.
      expect(code).toContain('"values":"\\"); while(true) {} var x = (\\""')
      expect(code).not.toMatch(/var args = [^;]*";/)
    })

    it('escapes newlines in injected args', async () => {
      const { code } = await callTool('write_range', { range: 'A1', values: 'line1\nline2' }, {})
      expect(code).toContain('"values":"line1\\nline2"')
    })
  })

  describe('inspect_workbook', () => {
    it('returns sheets, tables, names and pivots', async () => {
      const overview = {
        sheets: [{ name: 'Sheet1', position: 0, visibility: 'Visible', usedRange: 'Sheet1!A1:D7' }],
        tables: [{ name: 'SalesTable', worksheet: 'Sheet1', range: 'Sheet1!A3:D7' }],
        names: [{ name: 'TaxRate', value: '=Sheet1!$B$1', scope: 'Workbook' }],
        pivotTables: [],
      }
      const { code, result } = await callTool('inspect_workbook', { workbookId: 'book.xlsx' }, overview)
      expect(code).toContain('getUsedRangeOrNullObject')
      expect(code).toContain('pivotTables')
      expect(JSON.parse(result.text)).toEqual(overview)
    })
  })

  describe('format_range', () => {
    it('applies font, fill, alignment and borders', async () => {
      const { code } = await callTool(
        'format_range',
        { range: 'A1:D3', bold: true, fillColor: '#FFFF00', borders: 'all', workbookId: 'book.xlsx' },
        { formatted: 'A1:D3' },
      )
      expect(code).toContain('f.font.bold = args.bold')
      expect(code).toContain('f.fill.color = args.fillColor')
      expect(code).toContain('InsideHorizontal')
      expect(code).toContain('b.style = style')
    })

    it('clears borders with borders=none', async () => {
      const { code } = await callTool('format_range', { range: 'A1', borders: 'none' }, {})
      expect(code).toContain('style = "None"')
    })
  })

  describe('worksheet management', () => {
    it('add_worksheet injects name and position', async () => {
      const { code } = await callTool('add_worksheet', { name: 'Report', position: 2 }, { name: 'Report', position: 2 })
      expect(code).toContain('worksheets.add(args.name)')
      expect(code).toContain('sheet.position = args.position')
    })

    it('update_worksheet renames and moves', async () => {
      const { code } = await callTool(
        'update_worksheet',
        { sheetName: 'Sheet1', newName: 'Data', position: 1, visibility: 'Hidden' },
        { name: 'Data', position: 1, visibility: 'Hidden' },
      )
      expect(code).toContain('sheet.name = args.newName')
      expect(code).toContain('sheet.visibility = args.visibility')
    })

    it('delete_worksheet deletes by name', async () => {
      const { code } = await callTool('delete_worksheet', { sheetName: 'Old' }, { deleted: 'Old' })
      expect(code).toContain('getItem(args.sheetName).delete()')
    })
  })

  describe('sort_range', () => {
    it('maps fields to sort keys', async () => {
      const { code } = await callTool(
        'sort_range',
        { range: 'A2:D100', fields: [{ column: 0, ascending: false }], hasHeaders: true },
        { sorted: 'A2:D100' },
      )
      expect(code).toContain('{ key: f.column, ascending: f.ascending !== false }')
      expect(code).toContain('apply(fields, args.hasHeaders === true)')
    })
  })

  describe('find_in_range', () => {
    it('searches the used range by default and collects all matches', async () => {
      const { code } = await callTool('find_in_range', { text: 'needle' }, { found: true, address: 'Sheet1!B4' })
      expect(code).toContain('sheet.getUsedRange()')
      expect(code).toContain('range.find(args.text, criteria)')
      expect(code).toContain('range.findNext(current, criteria)')
    })

    it('returns only the first match when all=false', async () => {
      const { code } = await callTool('find_in_range', { text: 'needle', all: false }, { found: true, address: 'B4' })
      expect(code).toContain('args.all === false')
    })
  })

  describe('add_table', () => {
    it('creates a table with name and style', async () => {
      const { code } = await callTool(
        'add_table',
        { range: 'A3:D7', name: 'SalesTable', style: 'TableStyleMedium2' },
        { name: 'SalesTable', range: 'Sheet1!A3:D7' },
      )
      expect(code).toContain('sheet.tables.add(args.range')
      expect(code).toContain('table.style = args.style')
    })
  })

  describe('add_chart', () => {
    it('creates a chart with type and title', async () => {
      const { code } = await callTool(
        'add_chart',
        { range: 'A3:C7', chartType: 'Pie', title: 'Share' },
        { name: 'Chart 1', chartType: 'Pie' },
      )
      expect(code).toContain('sheet.charts.add(args.chartType || "ColumnClustered"')
      expect(code).toContain('chart.title.text = args.title')
    })
  })

  describe('add_conditional_format', () => {
    it('builds a cellValue rule with fill', async () => {
      const { code } = await callTool(
        'add_conditional_format',
        { range: 'D4:D7', ruleType: 'cellValue', operator: 'GreaterThan', formula1: '200', fillColor: '#FF0000' },
        { added: 'CellValue', priority: 1 },
      )
      expect(code).toContain('cellValue: { operator: args.operator')
      expect(code).toContain('fill: { color: args.fillColor')
    })

    it('builds a colorScale rule', async () => {
      const { code } = await callTool(
        'add_conditional_format',
        { range: 'B4:B7', ruleType: 'colorScale', minColor: '#F8696B', maxColor: '#63BE7B' },
        { added: 'ColorScale', priority: 1 },
      )
      expect(code).toContain('colorScale: {')
      expect(code).toContain('Minimum')
    })

    it('rejects cellValue without operator', async () => {
      const r = await callToolError('add_conditional_format', { range: 'A1', ruleType: 'cellValue' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('require an operator')
    })
  })

  describe('add_name', () => {
    it('defines a named range', async () => {
      const { code } = await callTool(
        'add_name',
        { name: 'TaxRate', ref: '=Sheet1!$B$1' },
        { name: 'TaxRate', value: '=Sheet1!$B$1' },
      )
      expect(code).toContain('names.add(args.name, args.ref')
    })
  })

  describe('set_data_validation', () => {
    it('builds a list rule', async () => {
      const { code } = await callTool(
        'set_data_validation',
        { range: 'E4:E7', type: 'List', list: 'Low,Medium,High' },
        { validated: 'E4:E7', type: 'List' },
      )
      expect(code).toContain('list: { inCellDropDown: true, source: args.list }')
    })

    it('builds a wholeNumber rule', async () => {
      const { code } = await callTool(
        'set_data_validation',
        { range: 'E4:E7', type: 'WholeNumber', operator: 'GreaterThan', formula1: '0' },
        { validated: 'E4:E7', type: 'WholeNumber' },
      )
      expect(code).toContain('inner[key] = { operator: args.operator')
      expect(code).toContain('ignoreBlanks')
    })

    it('rejects numeric rules without operator', async () => {
      const r = await callToolError('set_data_validation', { range: 'A1', type: 'WholeNumber' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('requires an operator')
    })
  })

  describe('calculate_workbook', () => {
    it('forces a recalculation', async () => {
      const { code } = await callTool('calculate_workbook', { type: 'Full' }, { calculated: 'Full' })
      expect(code).toContain('workbook.application.calculate(args.type || "Recalculate")')
    })
  })

  describe('insert_or_delete_range', () => {
    it('inserts with shift Down by default', async () => {
      const { code } = await callTool('insert_or_delete_range', { range: 'A5:D5', action: 'insert' }, {})
      expect(code).toContain('range.insert(args.shift || "Down")')
    })

    it('deletes with shift Up by default', async () => {
      const { code } = await callTool('insert_or_delete_range', { range: 'A5:D5', action: 'delete' }, {})
      expect(code).toContain('range.delete(args.shift || "Up")')
    })
  })

  describe('set_merge', () => {
    it('merges and unmerges', async () => {
      const merge = await callTool('set_merge', { range: 'A1:D1', action: 'merge' }, {})
      expect(merge.code).toContain('range.merge(args.across === true)')
      const unmerge = await callTool('set_merge', { range: 'A1:D1', action: 'unmerge' }, {})
      expect(unmerge.code).toContain('range.unmerge()')
    })
  })

  describe('autofit_range', () => {
    it('autofits columns and rows', async () => {
      const { code } = await callTool('autofit_range', { range: 'A:F' }, {})
      expect(code).toContain('autofitColumns()')
      expect(code).toContain('autofitRows()')
    })
  })

  describe('add_comment', () => {
    it('adds a threaded comment at a cell address', async () => {
      const { code } = await callTool('add_comment', { cellAddress: 'Sheet1!A1', text: 'Check this' }, {})
      expect(code).toContain('comments.add(args.cellAddress, args.text)')
    })
  })

  describe('refresh_pivottables', () => {
    it('refreshes all pivots', async () => {
      const { code } = await callTool('refresh_pivottables', {}, { refreshed: true })
      expect(code).toContain('pivotTables.refreshAll()')
    })
  })

  describe('set_worksheet_protection', () => {
    it('protects with a password', async () => {
      const { code } = await callTool(
        'set_worksheet_protection',
        { sheetName: 'Sheet1', protect: true, password: 'secret' },
        { sheet: 'Sheet1', protected: true },
      )
      expect(code).toContain('protection.protect(args.password)')
    })

    it('unprotects', async () => {
      const { code } = await callTool('set_worksheet_protection', { sheetName: 'Sheet1', protect: false }, {})
      expect(code).toContain('protection.unprotect(args.password)')
    })
  })

  describe('set_dimensions', () => {
    it('sets column width and row height', async () => {
      const { code } = await callTool('set_dimensions', { range: 'A:A', columnWidth: 120 }, {})
      expect(code).toContain('range.format.columnWidth = args.columnWidth')
    })
  })

  describe('read_comments', () => {
    it('loads comment fields and replies', async () => {
      const { code, result } = await callTool(
        'read_threaded_comments',
        {},
        {
          comments: [{ id: 'c1', content: 'Check this', author: 'A', resolved: false, replies: [] }],
        },
      )
      expect(code).toContain('context.workbook.comments')
      expect(code).toContain('replies.load("id,content,authorName,creationDate")')
      expect(JSON.parse(result.text).comments[0].id).toBe('c1')
    })
  })

  describe('reply_to_comment', () => {
    it('adds a reply to the comment id', async () => {
      const { code, result } = await callTool('reply_to_comment', { commentId: 'c1', text: 'Done' }, { replied: 'c1' })
      expect(code).toContain('comments.getItem(args.commentId)')
      expect(code).toContain('replies.add(args.text)')
      expect(JSON.parse(result.text).replied).toBe('c1')
    })

    it('injects the reply text as JSON, not interpolated', async () => {
      const { code } = await callTool('reply_to_comment', { commentId: 'c1', text: '"); evil()' }, {})
      expect(code).toContain('var args = {"commentId":"c1","text":"\\"); evil()"}')
    })
  })

  describe('resolve_comment', () => {
    it('sets the resolved flag and reads it back', async () => {
      const { code, result } = await callTool(
        'resolve_comment',
        { commentId: 'c1', resolved: true },
        { commentId: 'c1', resolved: true },
      )
      expect(code).toContain('comment.resolved = args.resolved')
      expect(JSON.parse(result.text).resolved).toBe(true)
    })
  })

  describe('set_hyperlink', () => {
    it('sets an object hyperlink with text and tooltip', async () => {
      const { code, result } = await callTool(
        'set_hyperlink',
        { cellAddress: 'Sheet1!B2', address: 'https://example.com', textToDisplay: 'Example' },
        { cell: 'Sheet1!B2', hyperlink: { address: 'https://example.com/' } },
      )
      expect(code).toContain('range.hyperlink = link')
      expect(code).toContain('link.textToDisplay = args.textToDisplay')
      expect(JSON.parse(result.text).hyperlink.address).toBe('https://example.com/')
    })

    it('clears the hyperlink when address is null', async () => {
      const { code } = await callTool('set_hyperlink', { cellAddress: 'Sheet1!B2', address: null }, {})
      expect(code).toContain('range.hyperlink = null')
    })

    it('splits the sheet-qualified address', async () => {
      const { code } = await callTool('set_hyperlink', { cellAddress: 'Q3 Data!B2', address: 'https://x.dev' }, {})
      expect(code).toContain('getItem(args.cellAddress.split("!")[0])')
    })
  })

  describe('freeze_panes', () => {
    it('freezes at the given cell', async () => {
      const { code, result } = await callTool(
        'freeze_panes',
        { sheetName: 'Sheet1', freezeAt: 'B2' },
        { sheet: 'Sheet1', frozenAt: 'B2' },
      )
      expect(code).toContain('freezePanes.freezeAt(sheet.getRange(args.freezeAt))')
      expect(JSON.parse(result.text).frozenAt).toBe('B2')
    })

    it('unfreezes when freezeAt is null', async () => {
      const { code } = await callTool('freeze_panes', { sheetName: 'Sheet1', freezeAt: null }, {})
      expect(code).toContain('freezePanes.unfreeze()')
    })
  })

  describe('replace_in_sheet', () => {
    it('uses findAll and rewrites matching cells', async () => {
      const { code, result } = await callTool(
        'replace_in_sheet',
        { find: '2025', replace: '2026' },
        { replacedCells: 3, matches: 3 },
      )
      expect(code).toContain('sheet.findAll(args.find')
      expect(code).toContain('live[i].values = grid')
      expect(JSON.parse(result.text).replacedCells).toBe(3)
    })

    it('handles case-insensitive replacement without corrupting casing', async () => {
      const { code } = await callTool('replace_in_sheet', { find: 'q3', replace: 'Q4' }, {})
      expect(code).toContain('var hay = args.matchCase ? v : v.toLowerCase()')
      expect(code).toContain('hay.indexOf(needle, pos)')
    })

    it('injects find and replace as JSON', async () => {
      const { code } = await callTool('replace_in_sheet', { find: 'Q"1', replace: 'Q"2' }, {})
      expect(code).toContain('var args = {"find":"Q\\"1","replace":"Q\\"2"}')
    })
  })

  describe('clear_range', () => {
    it('clears the range', async () => {
      const { code, result } = await callTool('clear_range', { range: 'A1:C3' }, { cleared: 'A1:C3' })
      expect(code).toContain('sheet.getRange(args.range).clear()')
      expect(JSON.parse(result.text).cleared).toBe('A1:C3')
    })
  })

  describe('copy_range', () => {
    it('copies values only when requested', async () => {
      const { code, result } = await callTool(
        'copy_range',
        { source: 'A1:C3', target: 'E1', valuesOnly: true },
        { copied: 'A1:C3 -> E1', valuesOnly: true },
      )
      expect(code).toContain('dst.copyFrom(src, "Values")')
      expect(JSON.parse(result.text).valuesOnly).toBe(true)
    })

    it('copies everything by default', async () => {
      const { code } = await callTool('copy_range', { source: 'A1:C3', target: 'E1' }, {})
      expect(code).toContain('dst.copyFrom(src)')
    })
  })
})
