import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'

/**
 * Word tools for the o365 extension.
 *
 * Thin, typed wrappers over the Office.js WordApi, following the portable
 * subset of the live Word MCP projects (all MIT, design reference):
 * kennypeh85/word-mcp, HelloWorld-Open/word-mcp-server,
 * Han-ops-dev/word-mcp-server.
 *
 * Tool bodies run inside Word.run(async (context) => { ... }) in the add-in.
 * Arguments are injected as a JSON literal (`args`), never string-interpolated.
 * Paragraph and table indices are 0-based.
 */

/** Run an Office.js body with JSON-injected args against a document. */
async function execWord(
  pool: ConnectionPool,
  documentId: string | undefined,
  args: unknown,
  body: string,
): Promise<unknown> {
  const target = pool.resolveTarget(documentId)
  const code = `var args = ${JSON.stringify(args)};\n${body}`
  return pool.sendCommand('executeCode', { code }, target.ws)
}

/**
 * Resolve a paragraph by 0-based index. `var paragraph` is available to the
 * rest of the generated body after this prelude.
 */
const PARAGRAPH_LINE = `
var paragraphs = context.document.body.paragraphs;
paragraphs.load("text,style");
await context.sync();
if (args.index < 0 || args.index >= paragraphs.items.length) {
  throw new Error("Paragraph index " + args.index + " out of range (0-" + (paragraphs.items.length - 1) + ")");
}
var paragraph = paragraphs.items[args.index];
`

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
    "Execute arbitrary Office.js code inside the live Word document. The code runs inside Word.run(async (context) => { ... }) with 'context' available as a variable. Use 'await context.sync()' after loading properties. Return a value to get it back as the tool result. Common operations: read the body (context.document.body), insert paragraphs, search and replace, work with tables, read tracked changes.",
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

  server.tool(
    'inspect_document',
    'Overview of the live document: paragraph count, section count, table count, inline picture count, and a preview of the first paragraphs with their styles. Use this first to orient yourself in an unfamiliar document.',
    {
      previewParagraphs: z
        .number()
        .int()
        .min(0)
        .max(50)
        .optional()
        .describe('How many paragraphs to preview. Default 10.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var body = context.document.body;
var paragraphs = body.paragraphs;
paragraphs.load("text,style");
var sections = context.document.sections;
sections.load("items");
var tables = body.tables;
tables.load("rowCount,values");
var pictures = body.inlinePictures;
pictures.load("items");
await context.sync();
var previewCount = args.previewParagraphs === undefined ? 10 : args.previewParagraphs;
var preview = [];
for (var i = 0; i < Math.min(previewCount, paragraphs.items.length); i++) {
  preview.push({ index: i, style: paragraphs.items[i].style, text: paragraphs.items[i].text.slice(0, 120) });
}
return {
  paragraphCount: paragraphs.items.length,
  sectionCount: sections.items.length,
  tableCount: tables.items.length,
  tables: tables.items.map(function (t) {
    return { rowCount: t.rowCount, columnCount: t.values && t.values[0] ? t.values[0].length : 0 };
  }),
  inlinePictureCount: pictures.items.length,
  preview: preview,
};
`,
      ),
    ),
  )

  server.tool(
    'read_paragraphs',
    'Read paragraphs of the document with their styles, paginated by offset and limit. Returns total count plus the requested slice.',
    {
      offset: z.number().int().min(0).optional().describe('First paragraph index to return. Default 0.'),
      limit: z.number().int().min(1).max(200).optional().describe('Maximum paragraphs to return. Default 50.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var paragraphs = context.document.body.paragraphs;
paragraphs.load("text,style");
await context.sync();
var offset = args.offset === undefined ? 0 : args.offset;
var limit = args.limit === undefined ? 50 : args.limit;
var slice = [];
for (var i = offset; i < Math.min(offset + limit, paragraphs.items.length); i++) {
  slice.push({ index: i, style: paragraphs.items[i].style, text: paragraphs.items[i].text });
}
return { total: paragraphs.items.length, offset: offset, paragraphs: slice };
`,
      ),
    ),
  )

  server.tool(
    'get_paragraph',
    'Read one paragraph in full: text, style, alignment and font properties (name, size, bold, italic, color).',
    {
      index: z.number().int().min(0).describe('Zero-based paragraph index.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
${PARAGRAPH_LINE}
var range = paragraph.getRange();
range.load("text,style");
range.font.load("name,size,bold,italic,color");
paragraph.load("alignment");
await context.sync();
return {
  index: args.index,
  text: range.text,
  style: range.style,
  alignment: paragraph.alignment,
  font: { name: range.font.name, size: range.font.size, bold: range.font.bold, italic: range.font.italic, color: range.font.color },
};
`,
      ),
    ),
  )

  server.tool(
    'insert_paragraph',
    'Insert a new paragraph into the document, optionally with a style (e.g. "Heading 1", "Title", "Quote"). Insert at the start or end of the body, or before/after a given paragraph index. Returns the new paragraph count.',
    {
      text: z.string().describe('Paragraph text. Use empty string for a blank paragraph.'),
      style: z
        .string()
        .optional()
        .describe(
          'Style name, e.g. "Heading 1", "Heading 2", "Title", "Quote", "Normal". Built-in names are matched locale-independently.',
        ),
      location: z
        .enum(['start', 'end', 'beforeIndex', 'afterIndex'])
        .optional()
        .describe('Where to insert. Default "end".'),
      index: z.number().int().min(0).optional().describe('Paragraph index for location beforeIndex/afterIndex.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) => {
      if ((p.location === 'beforeIndex' || p.location === 'afterIndex') && p.index === undefined) {
        throw new Error(`location "${p.location}" requires an index.`)
      }
      return execWord(
        pool,
        p.documentId,
        p,
        `
var body = context.document.body;
var paragraphs = body.paragraphs;
paragraphs.load("text,style");
await context.sync();
var location = args.location || "end";
var range;
if (location === "start") {
  range = body.insertParagraph(args.text, "Start");
} else if (location === "beforeIndex" || location === "afterIndex") {
  if (args.index >= paragraphs.items.length) {
    throw new Error("Paragraph index " + args.index + " out of range (0-" + (paragraphs.items.length - 1) + ")");
  }
  range = paragraphs.items[args.index].getRange().insertParagraph(args.text, location === "beforeIndex" ? "Before" : "After");
} else {
  range = body.insertParagraph(args.text, "End");
}
if (args.style !== undefined) {
  var builtIn = { "normal": "Normal", "title": "Title", "subtitle": "Subtitle", "quote": "Quote", "intensequote": "IntenseQuote", "listparagraph": "ListParagraph", "noSpacing": "NoSpacing" };
  for (var h = 1; h <= 9; h++) builtIn["heading" + h] = "Heading" + h;
  var key = args.style.toLowerCase().replace(/\\s+/g, "");
  if (builtIn[key] !== undefined) {
    range.styleBuiltIn = builtIn[key];
  } else {
    range.style = args.style;
  }
}
paragraphs = body.paragraphs;
paragraphs.load("text");
await context.sync();
return { inserted: args.text.slice(0, 80), paragraphCount: paragraphs.items.length };
`,
      )
    }),
  )

  server.tool(
    'set_paragraph_text',
    'Replace the text of one paragraph, keeping the paragraph itself (and optionally its style).',
    {
      index: z.number().int().min(0).describe('Zero-based paragraph index.'),
      text: z.string().describe('New paragraph text.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
${PARAGRAPH_LINE}
paragraph.getRange().insertText(args.text, "Replace");
await context.sync();
return { index: args.index, text: args.text };
`,
      ),
    ),
  )

  server.tool(
    'delete_paragraph',
    'Delete one paragraph (text and paragraph mark) from the document.',
    {
      index: z.number().int().min(0).describe('Zero-based paragraph index.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
${PARAGRAPH_LINE}
paragraph.getRange("Whole").delete();
await context.sync();
return { deleted: args.index };
`,
      ),
    ),
  )

  server.tool(
    'format_paragraph',
    'Format one paragraph: style, alignment, and font (name, size, bold, italic, color, highlight color). Only provided properties are changed.',
    {
      index: z.number().int().min(0).describe('Zero-based paragraph index.'),
      style: z
        .string()
        .optional()
        .describe('Style name, e.g. "Heading 1". Built-in names matched locale-independently.'),
      alignment: z.enum(['Left', 'Centered', 'Right', 'Justified']).optional().describe('Paragraph alignment.'),
      fontName: z.string().optional().describe('Font family, e.g. "Calibri".'),
      fontSize: z.number().positive().optional().describe('Font size in points.'),
      bold: z.boolean().optional().describe('Bold text.'),
      italic: z.boolean().optional().describe('Italic text.'),
      color: z.string().optional().describe('Font color as #RRGGBB.'),
      highlightColor: z.string().optional().describe('Highlight color as #RRGGBB. Use "#FFFFFF" to clear.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
${PARAGRAPH_LINE}
if (args.alignment !== undefined) paragraph.alignment = args.alignment;
var range = paragraph.getRange();
if (args.style !== undefined) {
  var builtIn = { "normal": "Normal", "title": "Title", "subtitle": "Subtitle", "quote": "Quote", "intensequote": "IntenseQuote", "listparagraph": "ListParagraph", "nospacing": "NoSpacing" };
  for (var h = 1; h <= 9; h++) builtIn["heading" + h] = "Heading" + h;
  var key = args.style.toLowerCase().replace(/\\s+/g, "");
  if (builtIn[key] !== undefined) {
    range.styleBuiltIn = builtIn[key];
  } else {
    range.style = args.style;
  }
}
var font = range.font;
if (args.fontName !== undefined) font.name = args.fontName;
if (args.fontSize !== undefined) font.size = args.fontSize;
if (args.bold !== undefined) font.bold = args.bold;
if (args.italic !== undefined) font.italic = args.italic;
if (args.color !== undefined) font.color = args.color;
if (args.highlightColor !== undefined) font.highlightColor = args.highlightColor;
await context.sync();
return { formatted: args.index };
`,
      ),
    ),
  )

  server.tool(
    'search_document',
    'Search the document body for text. Returns up to 50 matches with a text snippet around each hit.',
    {
      text: z.string().describe('Text to find.'),
      matchCase: z.boolean().optional().describe('Case-sensitive search. Default false.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var results = context.document.body.search(args.text, { matchCase: args.matchCase === true, matchWholeWord: false });
results.load("text");
await context.sync();
var matches = [];
for (var i = 0; i < Math.min(50, results.items.length); i++) {
  matches.push(results.items[i].text.slice(0, 160));
}
return { total: results.items.length, matches: matches };
`,
      ),
    ),
  )

  server.tool(
    'replace_text',
    'Find and replace text throughout the document. Returns the number of replacements made.',
    {
      find: z.string().describe('Text to find.'),
      replace: z.string().describe('Replacement text.'),
      matchCase: z.boolean().optional().describe('Case-sensitive search. Default false.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var results = context.document.body.search(args.find, { matchCase: args.matchCase === true, matchWholeWord: false });
results.load("text");
await context.sync();
var count = results.items.length;
for (var i = 0; i < results.items.length; i++) {
  results.items[i].insertText(args.replace, "Replace");
}
await context.sync();
return { replaced: count };
`,
      ),
    ),
  )

  server.tool(
    'insert_table',
    'Insert a table at the end of the document, optionally pre-filled from a 2D array of strings. Returns row and column counts.',
    {
      rows: z.number().int().min(1).max(100).describe('Number of rows.'),
      columns: z.number().int().min(1).max(50).describe('Number of columns.'),
      values: z
        .array(z.array(z.string()))
        .optional()
        .describe('Optional 2D array of cell texts (rows x columns) to pre-fill the table.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) => {
      if (p.values && (p.values.length !== p.rows || p.values.some((row) => row.length !== p.columns))) {
        throw new Error(`values must be a ${p.rows}x${p.columns} grid.`)
      }
      return execWord(
        pool,
        p.documentId,
        p,
        `
var body = context.document.body;
var table = args.values !== undefined
  ? body.insertTable(args.rows, args.columns, "End", args.values)
  : body.insertTable(args.rows, args.columns, "End");
table.load("rowCount,values");
await context.sync();
return { rowCount: table.rowCount, columnCount: table.values && table.values[0] ? table.values[0].length : 0 };
`,
      )
    }),
  )

  server.tool(
    'read_table',
    'Read all cell texts of a table as a 2D array.',
    {
      tableIndex: z.number().int().min(0).describe('Zero-based table index.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var tables = context.document.body.tables;
tables.load("rowCount,values");
await context.sync();
if (args.tableIndex < 0 || args.tableIndex >= tables.items.length) {
  throw new Error("Table index " + args.tableIndex + " out of range (0-" + (tables.items.length - 1) + ")");
}
var table = tables.items[args.tableIndex];
var grid = (table.values || []).map(function (row) {
  return row.map(function (cell) { return String(cell).replace(/\\r/g, "").replace(/\\n/g, " ").trim(); });
});
return {
  tableIndex: args.tableIndex,
  rows: table.rowCount,
  columns: grid.length > 0 ? grid[0].length : 0,
  cells: grid,
};
`,
      ),
    ),
  )

  server.tool(
    'set_table_cell',
    'Write text into one table cell (replacing its content).',
    {
      tableIndex: z.number().int().min(0).describe('Zero-based table index.'),
      row: z.number().int().min(0).describe('Zero-based row index.'),
      column: z.number().int().min(0).describe('Zero-based column index.'),
      text: z.string().describe('New cell text.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var tables = context.document.body.tables;
tables.load("rowCount,values");
await context.sync();
if (args.tableIndex < 0 || args.tableIndex >= tables.items.length) {
  throw new Error("Table index " + args.tableIndex + " out of range (0-" + (tables.items.length - 1) + ")");
}
var table = tables.items[args.tableIndex];
var colCount = table.values && table.values[0] ? table.values[0].length : 0;
if (args.row >= table.rowCount || args.column >= colCount) {
  throw new Error("Cell (" + args.row + "," + args.column + ") out of range for table " + args.tableIndex + " (" + table.rowCount + "x" + colCount + ")");
}
table.getCell(args.row, args.column).body.insertText(args.text, "Replace");
await context.sync();
return { tableIndex: args.tableIndex, row: args.row, column: args.column, text: args.text };
`,
      ),
    ),
  )

  server.tool(
    'insert_page_break',
    'Insert a page break at the end of the document body.',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
context.document.body.insertBreak("Page", "End");
await context.sync();
return { inserted: "page break" };
`,
      ),
    ),
  )

  server.tool(
    'comment_on_paragraph',
    'Add a comment to a paragraph (anchored on its text).',
    {
      index: z.number().int().min(0).describe('Zero-based paragraph index to comment on.'),
      text: z.string().describe('Comment text.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
${PARAGRAPH_LINE}
paragraph.getRange().insertComment(args.text);
await context.sync();
return { commented: args.index, text: args.text };
`,
      ),
    ),
  )

  server.tool(
    'get_tracked_changes',
    'List tracked changes (revisions) in the document: type (insertion/deletion), author, date and a text snippet. Read-only; Office.js cannot create tracked changes.',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var changes = context.document.body.getRange().getTrackedChanges();
changes.load("author,date,type");
await context.sync();
var out = [];
for (var i = 0; i < Math.min(100, changes.items.length); i++) {
  var change = changes.items[i];
  var snippet = "";
  try {
    var r = change.getRange();
    r.load("text");
    await context.sync();
    snippet = r.text.slice(0, 120);
  } catch (_e) { /* some change types have no text range */ }
  out.push({ index: i, author: change.author, date: change.date, type: change.type, text: snippet });
}
return { total: changes.items.length, changes: out };
`,
      ),
    ),
  )

  server.tool(
    'get_selection',
    'Read the current user selection (cursor) in the document.',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var selection = context.document.getSelection();
selection.load("text");
await context.sync();
return { text: selection.text };
`,
      ),
    ),
  )

  server.tool(
    'save_document',
    'Save the document. If it has never been saved, Word applies its default save behavior (may prompt).',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
context.document.save();
await context.sync();
return { saved: true };
`,
      ),
    ),
  )

  server.tool(
    'list_content_controls',
    'List all content controls (template fields) in the document: id, title, tag, type, current text and placeholder. Use set_content_control_text with the id to fill them.',
    {
      tag: z.string().optional().describe('Only return content controls with this tag.'),
      title: z.string().optional().describe('Only return content controls with this title.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var controls = context.document.contentControls;
controls.load("id,title,tag,type,text,placeholderText,appearance");
await context.sync();
var items = controls.items.map(function (c) {
  return { id: c.id, title: c.title, tag: c.tag, type: c.type, text: c.text, placeholder: c.placeholderText, appearance: c.appearance };
});
if (args.tag !== undefined) items = items.filter(function (c) { return c.tag === args.tag; });
if (args.title !== undefined) items = items.filter(function (c) { return c.title === args.title; });
return { count: items.length, contentControls: items };
`,
      ),
    ),
  )

  server.tool(
    'set_content_control_text',
    'Replace the text of a content control (template field) identified by its id from list_content_controls. Optionally update its title and tag.',
    {
      id: z.number().int().describe('Content control id from list_content_controls.'),
      text: z.string().describe('New text for the content control.'),
      title: z.string().optional().describe('New title for the content control.'),
      tag: z.string().optional().describe('New tag for the content control.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var control = context.document.contentControls.getByIdOrNullObject(args.id);
control.load("isNullObject,id");
await context.sync();
if (control.isNullObject) {
  throw new Error("Content control " + args.id + " not found");
}
control.insertText(args.text, "Replace");
if (args.title !== undefined) control.title = args.title;
if (args.tag !== undefined) control.tag = args.tag;
await context.sync();
control.load("text,title,tag");
await context.sync();
return { id: control.id, text: control.text, title: control.title, tag: control.tag };
`,
      ),
    ),
  )

  server.tool(
    'read_comments',
    'Read all comments in the document body: id, author, text, resolved state, date, and replies. Note: adding replies is not supported by the current Word.js API set.',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var comments = context.document.body.getRange().getComments();
comments.load("id,author,content,resolved,createdDate,replies");
await context.sync();
for (var i = 0; i < comments.items.length; i++) {
  comments.items[i].replies.load("id,content,author");
}
await context.sync();
return {
  comments: comments.items.map(function (c) {
    return {
      id: c.id,
      author: c.author,
      content: c.content,
      resolved: c.resolved,
      created: c.createdDate,
      replies: c.replies.items.map(function (r) {
        return { id: r.id, author: r.author, content: r.content };
      }),
    };
  }),
};
`,
      ),
    ),
  )

  server.tool(
    'get_document_properties',
    'Read the document metadata: title, author, subject, keywords, comments, category, creation and last-save dates.',
    {
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var properties = context.document.properties;
properties.load("title,author,subject,keywords,comments,category,creationDate,lastSaveDate");
await context.sync();
return {
  title: properties.title,
  author: properties.author,
  subject: properties.subject,
  keywords: properties.keywords,
  comments: properties.comments,
  category: properties.category,
  created: properties.creationDate,
  lastSaved: properties.lastSaveDate,
};
`,
      ),
    ),
  )

  server.tool(
    'set_document_properties',
    'Set document metadata (title, author, subject, keywords, comments, category). Only provided fields are changed.',
    {
      title: z.string().optional().describe('Document title.'),
      author: z.string().optional().describe('Document author.'),
      subject: z.string().optional().describe('Document subject.'),
      keywords: z.string().optional().describe('Document keywords.'),
      comments: z.string().optional().describe('Document comments field.'),
      category: z.string().optional().describe('Document category.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var properties = context.document.properties;
if (args.title !== undefined) properties.title = args.title;
if (args.author !== undefined) properties.author = args.author;
if (args.subject !== undefined) properties.subject = args.subject;
if (args.keywords !== undefined) properties.keywords = args.keywords;
if (args.comments !== undefined) properties.comments = args.comments;
if (args.category !== undefined) properties.category = args.category;
await context.sync();
return { updated: true };
`,
      ),
    ),
  )

  server.tool(
    'edit_table_columns',
    'Add or delete columns of a table. Add: position "Before" inserts left of the first column, "After" right of the last column; optionally pre-fill with values (2D array, one column per entry). Delete: remove count columns starting at columnIndex (0-based).',
    {
      tableIndex: z.number().int().min(0).describe('Zero-based table index.'),
      action: z.enum(['add', 'delete']).describe('Add or delete columns.'),
      position: z
        .enum(['Before', 'After'])
        .optional()
        .describe('For add: "Before" inserts at the left edge, "After" at the right edge.'),
      count: z.number().int().min(1).optional().describe('Number of columns (default 1).'),
      values: z
        .array(z.array(z.string()))
        .optional()
        .describe('For add: cell texts as 2D array (rows x added columns).'),
      columnIndex: z.number().int().min(0).optional().describe('For delete: first column index to remove.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var tables = context.document.body.tables;
tables.load("items");
await context.sync();
if (args.tableIndex < 0 || args.tableIndex >= tables.items.length) {
  throw new Error("Table index " + args.tableIndex + " out of range (0-" + (tables.items.length - 1) + ")");
}
var table = tables.items[args.tableIndex];
if (args.action === "add") {
  var count = args.count === undefined ? 1 : args.count;
  if (args.values !== undefined) {
    table.addColumns(args.position, count, args.values);
  } else {
    table.addColumns(args.position, count);
  }
  await context.sync();
  return { tableIndex: args.tableIndex, action: "add", position: args.position, count: count };
}
var startIndex = args.columnIndex === undefined ? 0 : args.columnIndex;
var delCount = args.count === undefined ? 1 : args.count;
table.deleteColumns(startIndex, delCount);
await context.sync();
return { tableIndex: args.tableIndex, action: "delete", columnIndex: startIndex, count: delCount };
`,
      ),
    ),
  )

  server.tool(
    'set_table_cells',
    'Set the text of multiple table cells in one call. Rows and columns are 0-based indices into the table.',
    {
      tableIndex: z.number().int().min(0).describe('Zero-based table index.'),
      cells: z
        .array(
          z.object({
            row: z.number().int().min(0).describe('Zero-based row index.'),
            column: z.number().int().min(0).describe('Zero-based column index.'),
            text: z.string().describe('New cell text.'),
          }),
        )
        .min(1)
        .describe('Cells to update.'),
      documentId: z
        .string()
        .optional()
        .describe('Target document ID from list_documents. Optional when only one document is connected.'),
    },
    withTool(async (p) =>
      execWord(
        pool,
        p.documentId,
        p,
        `
var tables = context.document.body.tables;
tables.load("items");
await context.sync();
if (args.tableIndex < 0 || args.tableIndex >= tables.items.length) {
  throw new Error("Table index " + args.tableIndex + " out of range (0-" + (tables.items.length - 1) + ")");
}
var table = tables.items[args.tableIndex];
var results = [];
for (var i = 0; i < args.cells.length; i++) {
  var cell = table.getCellOrNullObject(args.cells[i].row, args.cells[i].column);
  cell.load("isNullObject");
  await context.sync();
  if (cell.isNullObject) {
    results.push({ row: args.cells[i].row, column: args.cells[i].column, error: "cell out of range" });
  } else {
    cell.body.insertText(args.cells[i].text, "Replace");
    results.push({ row: args.cells[i].row, column: args.cells[i].column, text: args.cells[i].text });
  }
}
await context.sync();
return { tableIndex: args.tableIndex, cells: results };
`,
      ),
    ),
  )
}
