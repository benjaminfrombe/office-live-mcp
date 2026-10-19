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

describe('Word tools', () => {
  let pool: ConnectionPool
  let ws: WebSocket

  beforeEach(() => {
    pool = new ConnectionPool(100)
    ws = mockWs()
    pool.add('doc.docx', {
      ws,
      ready: true,
      presentationId: 'doc.docx',
      filePath: '/path/doc.docx',
    })
  })

  async function callTool(
    name: string,
    args: Record<string, unknown>,
    responseData?: unknown,
  ): Promise<{
    code: string
    result: { text: string; isError?: boolean }
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

  async function callToolError(
    name: string,
    args: Record<string, unknown>,
  ): Promise<{ text: string; isError?: boolean }> {
    const { client } = await setupMcpClient(pool)
    const result = (await client.callTool({ name, arguments: args })) as {
      content: Array<{ text: string }>
      isError?: boolean
    }
    return { text: result.content[0].text, isError: result.isError }
  }

  // -------------------------------------------------------------------------

  describe('list_documents', () => {
    it('returns empty message with no connections', async () => {
      const emptyPool = new ConnectionPool(100)
      const { client } = await setupMcpClient(emptyPool)
      const result = (await client.callTool({ name: 'list_documents', arguments: {} })) as {
        content: Array<{ text: string }>
      }
      expect(result.content[0].text).toContain('No documents connected')
    })

    it('lists connected documents', async () => {
      const { client } = await setupMcpClient(pool)
      const result = (await client.callTool({ name: 'list_documents', arguments: {} })) as {
        content: Array<{ text: string }>
      }
      const parsed = JSON.parse(result.content[0].text)
      expect(parsed[0]).toEqual({ documentId: 'doc.docx', filePath: '/path/doc.docx', ready: true })
    })
  })

  describe('execute_wordjs', () => {
    it('passes raw code through and returns the result', async () => {
      const { code, result } = await callTool(
        'execute_wordjs',
        { code: 'return context.document.body.text', documentId: 'doc.docx' },
        'Hello',
      )
      expect(code).toContain('return context.document.body.text')
      expect(JSON.parse(result.text)).toBe('Hello')
    })

    it('rejects an unknown documentId', async () => {
      const r = await callToolError('execute_wordjs', { code: 'return 1', documentId: 'nope.docx' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('not found')
    })
  })

  describe('inspect_document', () => {
    it('returns counts and a paragraph preview', async () => {
      const overview = {
        paragraphCount: 3,
        sectionCount: 1,
        tableCount: 0,
        tables: [],
        inlinePictureCount: 0,
        preview: [{ index: 0, style: 'Heading 1', text: 'Title' }],
      }
      const { code, result } = await callTool('inspect_document', { documentId: 'doc.docx' }, overview)
      expect(code).toContain('body.paragraphs')
      expect(code).toContain('inlinePictures')
      expect(JSON.parse(result.text)).toEqual(overview)
    })
  })

  describe('read_paragraphs', () => {
    it('paginates paragraphs with styles', async () => {
      const { code, result } = await callTool(
        'read_paragraphs',
        { offset: 2, limit: 5 },
        { total: 10, offset: 2, paragraphs: [] },
      )
      expect(code).toContain('paragraphs.load("text,style")')
      expect(JSON.parse(result.text).offset).toBe(2)
    })
  })

  describe('get_paragraph', () => {
    it('loads text, style, alignment and font', async () => {
      const { code } = await callTool('get_paragraph', { index: 1 }, {})
      expect(code).toContain('out of range')
      expect(code).toContain('range.font.load("name,size,bold,italic,color")')
      expect(code).toContain('paragraph.alignment')
    })
  })

  describe('insert_paragraph', () => {
    it('inserts at the end with a built-in style', async () => {
      const { code } = await callTool('insert_paragraph', { text: 'New section', style: 'Heading 2' }, {})
      expect(code).toContain('body.insertParagraph(args.text, "End")')
      expect(code).toContain('builtIn["heading" + h] = "Heading" + h')
      expect(code).toContain('styleBuiltIn')
    })

    it('inserts after a paragraph index', async () => {
      const { code } = await callTool('insert_paragraph', { text: 'X', location: 'afterIndex', index: 3 }, {})
      expect(code).toContain('insertParagraph(args.text, location === "beforeIndex" ? "Before" : "After")')
    })

    it('uses range.style for custom styles', async () => {
      const { code } = await callTool('insert_paragraph', { text: 'X', style: 'My Custom Style' }, {})
      expect(code).toContain('range.style = args.style')
    })

    it('rejects beforeIndex/afterIndex without an index', async () => {
      const r = await callToolError('insert_paragraph', { text: 'X', location: 'afterIndex' })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('requires an index')
    })

    it('keeps hostile text inside the JSON literal', async () => {
      const hostile = '"); while(true) {} var x = ("'
      const { code } = await callTool('insert_paragraph', { text: hostile }, {})
      expect(code).toContain('"text":"\\"); while(true) {} var x = (\\""')
    })
  })

  describe('set_paragraph_text', () => {
    it('replaces paragraph text in place', async () => {
      const { code } = await callTool('set_paragraph_text', { index: 0, text: 'New text' }, {})
      expect(code).toContain('insertText(args.text, "Replace")')
    })
  })

  describe('delete_paragraph', () => {
    it('deletes the whole paragraph range', async () => {
      const { code } = await callTool('delete_paragraph', { index: 0 }, {})
      expect(code).toContain('getRange("Whole").delete()')
    })
  })

  describe('format_paragraph', () => {
    it('applies style, alignment and font', async () => {
      const { code } = await callTool(
        'format_paragraph',
        { index: 0, style: 'Heading 1', alignment: 'Centered', bold: true, highlightColor: '#FFFF00' },
        {},
      )
      expect(code).toContain('paragraph.alignment = args.alignment')
      expect(code).toContain('font.highlightColor = args.highlightColor')
      expect(code).toContain('font.bold = args.bold')
    })
  })

  describe('search_document', () => {
    it('searches the body and caps the results', async () => {
      const { code } = await callTool('search_document', { text: 'needle' }, { total: 2, matches: ['needle'] })
      expect(code).toContain('body.search(args.text')
      expect(code).toContain('Math.min(50, results.items.length)')
    })
  })

  describe('replace_text', () => {
    it('replaces every match', async () => {
      const { code, result } = await callTool('replace_text', { find: 'old', replace: 'new' }, { replaced: 3 })
      expect(code).toContain('results.items[i].insertText(args.replace, "Replace")')
      expect(JSON.parse(result.text).replaced).toBe(3)
    })
  })

  describe('insert_table', () => {
    it('inserts a pre-filled table', async () => {
      const { code } = await callTool(
        'insert_table',
        {
          rows: 2,
          columns: 2,
          values: [
            ['a', 'b'],
            ['c', 'd'],
          ],
        },
        { rowCount: 2, columnCount: 2 },
      )
      expect(code).toContain('body.insertTable(args.rows, args.columns, "End", args.values)')
    })

    it('inserts an empty table without values', async () => {
      const { code } = await callTool('insert_table', { rows: 3, columns: 3 }, { rowCount: 3, columnCount: 3 })
      expect(code).toContain('body.insertTable(args.rows, args.columns, "End")')
    })

    it('rejects a mismatched values grid', async () => {
      const r = await callToolError('insert_table', { rows: 2, columns: 2, values: [['a']] })
      expect(r.isError).toBe(true)
      expect(r.text).toContain('2x2 grid')
    })
  })

  describe('read_table', () => {
    it('reads the cell grid via table.values', async () => {
      const { code, result } = await callTool(
        'read_table',
        { tableIndex: 0 },
        { tableIndex: 0, rows: 1, columns: 2, cells: [['a', 'b']] },
      )
      expect(code).toContain('tables.load("rowCount,values")')
      expect(code).toContain('table.values')
      expect(JSON.parse(result.text).cells).toEqual([['a', 'b']])
    })
  })

  describe('set_table_cell', () => {
    it('writes one cell and validates bounds', async () => {
      const { code } = await callTool('set_table_cell', { tableIndex: 0, row: 1, column: 2, text: 'cell' }, {})
      expect(code).toContain('getCell(args.row, args.column).body.insertText(args.text, "Replace")')
      expect(code).toContain('out of range for table')
    })
  })

  describe('insert_page_break', () => {
    it('inserts a page break at the end', async () => {
      const { code } = await callTool('insert_page_break', {}, {})
      expect(code).toContain('insertBreak("Page", "End")')
    })
  })

  describe('comment_on_paragraph', () => {
    it('comments on a paragraph', async () => {
      const { code } = await callTool('comment_on_paragraph', { index: 0, text: 'Check this' }, {})
      expect(code).toContain('insertComment(args.text)')
    })
  })

  describe('get_tracked_changes', () => {
    it('lists revisions with author and type', async () => {
      const { code } = await callTool('get_tracked_changes', {}, { total: 0, changes: [] })
      expect(code).toContain('getTrackedChanges()')
      expect(code).toContain('changes.load("author,date,type")')
    })
  })

  describe('get_selection', () => {
    it('reads the current selection', async () => {
      const { code } = await callTool('get_selection', {}, { text: 'selected' })
      expect(code).toContain('getSelection()')
    })
  })

  describe('save_document', () => {
    it('saves the document', async () => {
      const { code } = await callTool('save_document', {}, { saved: true })
      expect(code).toContain('context.document.save()')
    })
  })
})
