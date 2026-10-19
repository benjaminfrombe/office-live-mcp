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
  registerTools(server, pool, () => 'test-session', () => 1)

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
  async function callTool(name: string, args: Record<string, unknown>, responseData?: unknown): Promise<{
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
        { range: 'A1:B2', values: [[1, 2], [3, 4]], formulas: [['=A1', '=B1'], ['=A2', '=B2']], workbookId: 'book.xlsx' },
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
    it('searches the used range by default with findAll', async () => {
      const { code } = await callTool('find_in_range', { text: 'needle' }, { found: true, address: 'Sheet1!B4' })
      expect(code).toContain('sheet.getUsedRange()')
      expect(code).toContain('findAllOrNullObject')
    })

    it('uses find when all=false', async () => {
      const { code } = await callTool('find_in_range', { text: 'needle', all: false }, { found: true, address: 'B4' })
      expect(code).toContain('range.find(args.text, criteria)')
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
      expect(code).toContain('{ type: "List", list: args.list')
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
      expect(code).toContain('workbook.calculate(args.type || "Recalculate")')
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
})
