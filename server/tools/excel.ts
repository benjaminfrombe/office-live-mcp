import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'

/**
 * Excel tools for the o365 extension.
 *
 * The bridge, connection pool and executeCode escape hatch are host-neutral:
 * an Excel add-in connecting to the same bridge registers as a connection
 * whose code runs inside Excel.run(). These tools are thin, typed wrappers
 * over the Office.js ExcelApi, following the tool surface of
 * sbroenne/mcp-server-excel (MIT) where it is portable to Office.js on macOS.
 *
 * Tool bodies run inside Excel.run(async (context) => { ... }) in the add-in.
 * Arguments are injected as a JSON literal (`args`), never string-interpolated,
 * so user data cannot break out of the JSON context.
 */

/** Run an Office.js body with JSON-injected args against a workbook. */
async function execExcel(
  pool: ConnectionPool,
  workbookId: string | undefined,
  args: unknown,
  body: string,
): Promise<unknown> {
  const target = pool.resolveTarget(workbookId)
  const code = `var args = ${JSON.stringify(args)};\n${body}`
  return pool.sendCommand('executeCode', { code }, target.ws)
}

/** Resolve a worksheet object reference line used by most sheet tools. */
const SHEET_LINE = `
var wbSheets = context.workbook.worksheets;
var sheet = args.sheetName ? wbSheets.getItem(args.sheetName) : wbSheets.getActiveWorksheet();
`

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

  server.tool(
    'inspect_workbook',
    'Overview of the live workbook: all worksheets (name, position, visibility, used-range address), all Excel tables (name, sheet, range), all named ranges, and all PivotTables (name, sheet). Use this first to orient yourself in an unfamiliar workbook.',
    {
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async ({ workbookId }) =>
      execExcel(
        pool,
        workbookId,
        {},
        `
var wb = context.workbook;
var sheets = wb.worksheets.load("name,position,visibility");
var tables = wb.tables.load("name,worksheet");
var names = wb.names.load("name,value,scope");
var pivots = wb.pivotTables.load("name,worksheet");
await context.sync();
var usedRanges = [];
for (var i = 0; i < sheets.items.length; i++) {
  var u = sheets.items[i].getUsedRangeOrNullObject(true);
  u.load("address,isNullObject");
  usedRanges.push(u);
}
var tableRanges = [];
for (var j = 0; j < tables.items.length; j++) {
  var r = tables.items[j].getRange();
  r.load("address");
  tableRanges.push(r);
}
await context.sync();
var sheetData = [];
for (var i = 0; i < sheets.items.length; i++) {
  sheetData.push({
    name: sheets.items[i].name,
    position: sheets.items[i].position,
    visibility: sheets.items[i].visibility,
    usedRange: usedRanges[i].isNullObject ? null : usedRanges[i].address,
  });
}
var tableData = [];
for (var j = 0; j < tables.items.length; j++) {
  tableData.push({ name: tables.items[j].name, worksheet: tables.items[j].worksheet.name, range: tableRanges[j].address });
}
return {
  sheets: sheetData,
  tables: tableData,
  names: names.items.map(function (n) { return { name: n.name, value: n.value, scope: n.scope }; }),
  pivotTables: pivots.items.map(function (p) { return { name: p.name, worksheet: p.worksheet.name }; }),
};
`,
      ),
    ),
  )

  server.tool(
    'read_range',
    'Read values, formulas and number formats from a range of the live workbook. Returns a 2D array (rows x columns) plus the range address and dimensions. Use "usedRange": true to read the whole used range of the sheet.',
    {
      range: z
        .string()
        .optional()
        .describe('A1 range address, e.g. "A1:C10" or "MyTable". Omit when usedRange is true.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      usedRange: z
        .boolean()
        .optional()
        .describe('Read the entire used range of the worksheet instead of a specific address.'),
      includeFormulas: z.boolean().optional().describe('Include the formulas grid (default true).'),
      includeNumberFormat: z.boolean().optional().describe('Include the number-format grid (default false).'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = args.usedRange ? sheet.getUsedRange() : sheet.getRange(args.range);
range.load("address,rowCount,columnCount,values");
if (args.includeFormulas !== false) range.load("formulas");
if (args.includeNumberFormat) range.load("numberFormat");
await context.sync();
var out = { address: range.address, rowCount: range.rowCount, columnCount: range.columnCount, values: range.values };
if (args.includeFormulas !== false) out.formulas = range.formulas;
if (args.includeNumberFormat) out.numberFormat = range.numberFormat;
return out;
`,
      ),
    ),
  )

  server.tool(
    'write_range',
    'Write values and/or formulas into a range of the live workbook. Accepts a 2D array (rows x columns) or a single scalar that fills the whole range. Formulas are set separately from values. Also supports setting the number format. Content in the target range is overwritten.',
    {
      range: z.string().describe('A1 range address, e.g. "A1:C10".'),
      values: z
        .union([
          z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
          z.string(),
          z.number(),
          z.boolean(),
          z.null(),
        ])
        .optional()
        .describe('2D array of values, or a single scalar broadcast across the range.'),
      formulas: z
        .union([z.array(z.array(z.string())), z.string()])
        .optional()
        .describe('2D array of formula strings (e.g. "=SUM(A1:A10)"), or a single formula broadcast across the range.'),
      numberFormat: z
        .union([z.array(z.array(z.string())), z.string()])
        .optional()
        .describe('Number format string (e.g. "0.00", "#,##0", "dd/mm/yyyy"), 2D array or single broadcast.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) => {
      if (p.values === undefined && p.formulas === undefined && p.numberFormat === undefined) {
        throw new Error('Provide at least one of: values, formulas, numberFormat.')
      }
      return execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
if (args.numberFormat !== undefined) range.numberFormat = args.numberFormat;
if (args.values !== undefined) range.values = args.values;
if (args.formulas !== undefined) range.formulas = args.formulas;
range.load("address,rowCount,columnCount");
await context.sync();
return { written: range.address, rowCount: range.rowCount, columnCount: range.columnCount };
`,
      )
    }),
  )

  server.tool(
    'format_range',
    'Format a range: font (name, size, bold, italic, underline, color), fill color, borders (all/outline with style, weight, color), alignment, wrap, and number format. Only the provided properties are changed.',
    {
      range: z.string().describe('A1 range address, e.g. "A1:C10".'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      fontName: z.string().optional().describe('Font family, e.g. "Calibri".'),
      fontSize: z.number().positive().optional().describe('Font size in points.'),
      bold: z.boolean().optional().describe('Bold text.'),
      italic: z.boolean().optional().describe('Italic text.'),
      underline: z.enum(['None', 'Single', 'Double']).optional().describe('Underline style.'),
      fontColor: z.string().optional().describe('Font color as #RRGGBB, e.g. "#FF0000".'),
      fillColor: z.string().optional().describe('Cell fill color as #RRGGBB, e.g. "#FFFF00".'),
      horizontalAlignment: z.enum(['Left', 'Center', 'Right', 'General', 'Justify', 'Fill', 'Distributed']).optional(),
      verticalAlignment: z.enum(['Top', 'Center', 'Bottom', 'Justify', 'Distributed']).optional(),
      wrapText: z.boolean().optional().describe('Wrap text within cells.'),
      numberFormat: z.string().optional().describe('Number format string, e.g. "0.00%".'),
      borders: z.enum(['all', 'outline', 'none']).optional().describe('Which borders to apply. "none" clears borders.'),
      borderStyle: z
        .enum(['Continuous', 'Dash', 'DashDot', 'DashDotDot', 'Dot', 'Double', 'SlantDashDot'])
        .optional()
        .describe('Border line style. Default "Continuous".'),
      borderColor: z.string().optional().describe('Border color as #RRGGBB. Default "#000000".'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
var f = range.format;
if (args.fontName !== undefined) f.font.name = args.fontName;
if (args.fontSize !== undefined) f.font.size = args.fontSize;
if (args.bold !== undefined) f.font.bold = args.bold;
if (args.italic !== undefined) f.font.italic = args.italic;
if (args.underline !== undefined) f.font.underline = args.underline;
if (args.fontColor !== undefined) f.font.color = args.fontColor;
if (args.fillColor !== undefined) f.fill.color = args.fillColor;
if (args.horizontalAlignment !== undefined) f.horizontalAlignment = args.horizontalAlignment;
if (args.verticalAlignment !== undefined) f.verticalAlignment = args.verticalAlignment;
if (args.wrapText !== undefined) f.wrapText = args.wrapText;
if (args.numberFormat !== undefined) range.numberFormat = args.numberFormat;
if (args.borders !== undefined) {
  var style = args.borderStyle || "Continuous";
  var color = args.borderColor || "#000000";
  var edges = args.borders === "all"
    ? ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]
    : ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  if (args.borders === "none") style = "None";
  for (var i = 0; i < edges.length; i++) {
    var b = f.borders.getItem(edges[i]);
    b.style = style;
    if (args.borders !== "none") b.color = color;
  }
}
await context.sync();
return { formatted: args.range };
`,
      ),
    ),
  )

  server.tool(
    'set_dimensions',
    'Set column widths and/or row heights for a range, in points. Optionally autofit columns or rows to their content.',
    {
      range: z.string().describe('A1 range address, e.g. "A:A" or "1:10" or "A1:C10".'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      columnWidth: z.number().optional().describe('Column width in points.'),
      rowHeight: z.number().optional().describe('Row height in points.'),
      autofitColumns: z.boolean().optional().describe('Autofit column widths to content.'),
      autofitRows: z.boolean().optional().describe('Autofit row heights to content.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
if (args.columnWidth !== undefined) range.format.columnWidth = args.columnWidth;
if (args.rowHeight !== undefined) range.format.rowHeight = args.rowHeight;
if (args.autofitColumns) range.format.autofitColumns();
if (args.autofitRows) range.format.autofitRows();
await context.sync();
return { resized: args.range };
`,
      ),
    ),
  )

  server.tool(
    'add_worksheet',
    'Add a new worksheet to the workbook, optionally at a specific position, and activate it. Returns the new sheet name and position.',
    {
      name: z.string().optional().describe('Name for the new worksheet. Auto-generated when omitted.'),
      position: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('Zero-based position to insert at. Default: append at end.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var sheet = args.name !== undefined
  ? context.workbook.worksheets.add(args.name)
  : context.workbook.worksheets.add();
if (args.position !== undefined) sheet.position = args.position;
sheet.activate();
sheet.load("name,position");
await context.sync();
return { name: sheet.name, position: sheet.position };
`,
      ),
    ),
  )

  server.tool(
    'update_worksheet',
    'Update an existing worksheet: rename it, move it to a position, change visibility, set the tab color, or activate it. Only provided properties are changed.',
    {
      sheetName: z.string().describe('Current name of the worksheet to update.'),
      newName: z.string().optional().describe('New name for the worksheet.'),
      position: z.number().int().min(0).optional().describe('Zero-based position to move the sheet to.'),
      visibility: z.enum(['Visible', 'Hidden', 'VeryHidden']).optional().describe('Sheet visibility.'),
      tabColor: z.string().optional().describe('Tab color as #RRGGBB. Use "" to clear.'),
      activate: z.boolean().optional().describe('Make this the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var sheet = context.workbook.worksheets.getItem(args.sheetName);
if (args.newName !== undefined) sheet.name = args.newName;
if (args.position !== undefined) sheet.position = args.position;
if (args.visibility !== undefined) sheet.visibility = args.visibility;
if (args.tabColor !== undefined) sheet.tabColor = args.tabColor === "" ? null : args.tabColor;
if (args.activate) sheet.activate();
sheet.load("name,position,visibility");
await context.sync();
return { name: sheet.name, position: sheet.position, visibility: sheet.visibility };
`,
      ),
    ),
  )

  server.tool(
    'delete_worksheet',
    'Delete a worksheet from the workbook. This removes its contents permanently and can break formulas referencing it.',
    {
      sheetName: z.string().describe('Name of the worksheet to delete.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
context.workbook.worksheets.getItem(args.sheetName).delete();
await context.sync();
return { deleted: args.sheetName };
`,
      ),
    ),
  )

  server.tool(
    'sort_range',
    'Sort a range by one or more columns. Fields are given as zero-based column offsets within the range.',
    {
      range: z.string().describe('A1 range address to sort, e.g. "A2:D100" (include or exclude headers as needed).'),
      fields: z
        .array(
          z.object({
            column: z.number().int().min(0).describe('Zero-based column offset within the range.'),
            ascending: z.boolean().optional().describe('Sort direction. Default true.'),
          }),
        )
        .min(1)
        .describe('Sort fields, applied in order.'),
      hasHeaders: z
        .boolean()
        .optional()
        .describe('True if the first row is a header row that should not be sorted. Default false.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
var fields = args.fields.map(function (f) { return { key: f.column, ascending: f.ascending !== false }; });
range.sort.apply(fields, args.hasHeaders === true);
await context.sync();
return { sorted: args.range, fields: fields.length };
`,
      ),
    ),
  )

  server.tool(
    'find_in_range',
    'Find text in a range or the whole used range. Returns the address of the first match, or all matches as a (possibly multi-area) address string.',
    {
      text: z.string().describe('Text to find.'),
      range: z.string().optional().describe('A1 range address to search. Defaults to the used range.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      completeMatch: z.boolean().optional().describe('Match the entire cell contents. Default false.'),
      matchCase: z.boolean().optional().describe('Case-sensitive search. Default false.'),
      all: z.boolean().optional().describe('Return all matches instead of only the first. Default true.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
var criteria = { completeMatch: args.completeMatch === true, matchCase: args.matchCase === true };
var matches = [];
try {
  var current = range.find(args.text, criteria);
  while (matches.length < 500) {
    current.load("address");
    await context.sync();
    if (matches.indexOf(current.address) !== -1) break;
    matches.push(current.address);
    current = range.findNext(current, criteria);
  }
} catch (_e) {
  // find throws when there is no match at all
}
if (args.all === false) {
  return matches.length === 0 ? { found: false } : { found: true, address: matches[0] };
}
return matches.length === 0
  ? { found: false }
  : { found: true, address: matches.join(","), count: matches.length };
`,
      ),
    ),
  )

  server.tool(
    'add_table',
    'Create an Excel table (ListObject) over a range. Returns the table name and range.',
    {
      range: z.string().describe('A1 range address including headers, e.g. "A1:D20".'),
      name: z.string().optional().describe('Table name. Auto-generated when omitted.'),
      hasHeaders: z.boolean().optional().describe('True if the first row contains headers. Default true.'),
      style: z.string().optional().describe('Built-in table style name, e.g. "TableStyleMedium2".'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var table = sheet.tables.add(args.range, args.hasHeaders !== false);
if (args.name !== undefined) table.name = args.name;
if (args.style !== undefined) table.style = args.style;
table.load("name,style");
var tableRange = table.getRange();
tableRange.load("address");
await context.sync();
return { name: table.name, style: table.style, range: tableRange.address };
`,
      ),
    ),
  )

  server.tool(
    'add_chart',
    'Create a chart from a range. Chart types: ColumnClustered, ColumnStacked, BarClustered, BarStacked, Line, LineMarkers, Pie, Doughnut, Area, XYScatterLines. Returns the chart name and title.',
    {
      range: z.string().describe('A1 range address with the data (and labels) to plot.'),
      chartType: z
        .enum([
          'ColumnClustered',
          'ColumnStacked',
          'ColumnStacked100',
          'BarClustered',
          'BarStacked',
          'BarStacked100',
          'Line',
          'LineMarkers',
          'Pie',
          'PieStacked',
          'Doughnut',
          'Area',
          'AreaStacked',
          'XYScatterLines',
        ])
        .optional()
        .describe('Chart type. Default ColumnClustered.'),
      title: z.string().optional().describe('Chart title. Empty string removes the title.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var dataRange = sheet.getRange(args.range);
var chart = sheet.charts.add(args.chartType || "ColumnClustered", dataRange, "Auto");
if (args.title !== undefined) chart.title.text = args.title;
chart.load("name");
if (args.title !== undefined) chart.title.load("text");
await context.sync();
var out = { name: chart.name, chartType: args.chartType || "ColumnClustered" };
if (args.title !== undefined) out.title = chart.title.text;
return out;
`,
      ),
    ),
  )

  server.tool(
    'add_conditional_format',
    'Add a conditional formatting rule to a range. Supports cellValue (compare with a formula and fill on match), colorScale (min/max gradient), dataBar, and iconSet rules.',
    {
      range: z.string().describe('A1 range address.'),
      ruleType: z.enum(['cellValue', 'colorScale', 'dataBar', 'iconSet']).describe('Type of rule.'),
      operator: z
        .enum([
          'GreaterThan',
          'GreaterThanOrEqual',
          'LessThan',
          'LessThanOrEqual',
          'Between',
          'NotBetween',
          'EqualTo',
          'NotEqualTo',
        ])
        .optional()
        .describe('cellValue: comparison operator.'),
      formula1: z.string().optional().describe('cellValue: comparison value/formula, e.g. "0" or "=TODAY()".'),
      formula2: z.string().optional().describe('cellValue: second bound for Between/NotBetween.'),
      fillColor: z.string().optional().describe('cellValue: fill color on match, as #RRGGBB. Default "#FFC7CE".'),
      minColor: z.string().optional().describe('colorScale: minimum color. Default "#F8696B".'),
      maxColor: z.string().optional().describe('colorScale: maximum color. Default "#63BE7B".'),
      iconSetStyle: z
        .enum(['ThreeTrafficLights1', 'ThreeSymbols', 'ThreeArrows', 'FourArrows', 'FiveArrows', 'ThreeStars'])
        .optional()
        .describe('iconSet: icon set style. Default ThreeTrafficLights1.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) => {
      if (p.ruleType === 'cellValue' && !p.operator) {
        throw new Error('cellValue rules require an operator.')
      }
      return execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
var rule;
if (args.ruleType === "cellValue") {
  rule = {
    cellValue: { operator: args.operator, formula1: args.formula1, formula2: args.formula2 },
    format: { fill: { color: args.fillColor || "#FFC7CE" } },
  };
} else if (args.ruleType === "colorScale") {
  rule = {
    colorScale: {
      criteria: {
        min: { type: "Minimum", color: args.minColor || "#F8696B" },
        max: { type: "Maximum", color: args.maxColor || "#63BE7B" },
      },
    },
  };
} else if (args.ruleType === "dataBar") {
  rule = { dataBar: { showValue: true } };
} else {
  rule = { iconSet: { style: args.iconSetStyle || "ThreeTrafficLights1", reverseOrder: false, showValue: true } };
}
var cf = range.conditionalFormats.add(args.ruleType, rule);
cf.load("type,priority");
await context.sync();
return { added: cf.type, priority: cf.priority, range: args.range };
`,
      )
    }),
  )

  server.tool(
    'add_name',
    'Define a named range (named formula) in the workbook, e.g. name "TaxRate" for "=Sheet1!$B$1".',
    {
      name: z.string().describe('Name to define, e.g. "TaxRate".'),
      ref: z.string().describe('Reference formula, e.g. "=Sheet1!$A$1:$D$10".'),
      comment: z.string().optional().describe('Optional description.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var named = context.workbook.names.add(args.name, args.ref, args.comment);
named.load("name,value");
await context.sync();
return { name: named.name, value: named.value };
`,
      ),
    ),
  )

  server.tool(
    'set_data_validation',
    'Set a data validation rule on a range. Supports list, wholeNumber, decimal, textLength, and date rules.',
    {
      range: z.string().describe('A1 range address.'),
      type: z.enum(['List', 'WholeNumber', 'Decimal', 'TextLength', 'Date']).describe('Validation type.'),
      list: z.string().optional().describe('List type: comma-separated values, e.g. "Low,Medium,High".'),
      operator: z
        .enum([
          'GreaterThan',
          'GreaterThanOrEqual',
          'LessThan',
          'LessThanOrEqual',
          'Between',
          'NotBetween',
          'EqualTo',
          'NotEqualTo',
        ])
        .optional()
        .describe('Numeric/date/textLength: comparison operator.'),
      formula1: z.string().optional().describe('Numeric/date/textLength: bound value or formula.'),
      formula2: z.string().optional().describe('Second bound for Between/NotBetween.'),
      allowBlank: z.boolean().optional().describe('Allow blank cells. Default true.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) => {
      if (p.type !== 'List' && !p.operator) {
        throw new Error(`type "${p.type}" requires an operator.`)
      }
      return execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
var rule;
if (args.type === "List") {
  rule = { list: { inCellDropDown: true, source: args.list } };
} else {
  var key = args.type === "WholeNumber" ? "wholeNumber" : args.type === "Decimal" ? "decimal" : args.type === "Date" ? "date" : "textLength";
  var inner = {};
  inner[key] = { operator: args.operator, formula1: args.formula1, formula2: args.formula2 };
  rule = inner;
}
range.dataValidation.rule = rule;
range.dataValidation.ignoreBlanks = args.allowBlank !== false;
await context.sync();
return { validated: args.range, type: args.type };
`,
      )
    }),
  )

  server.tool(
    'calculate_workbook',
    'Force Excel to recalculate the workbook. Use after writes when you need formula results to be current before reading them back.',
    {
      type: z
        .enum(['Full', 'FullRebuild', 'Recalculate'])
        .optional()
        .describe('Recalculation type. Default "Recalculate".'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
context.workbook.application.calculate(args.type || "Recalculate");
await context.sync();
return { calculated: args.type || "Recalculate" };
`,
      ),
    ),
  )

  server.tool(
    'insert_or_delete_range',
    'Insert or delete a range, shifting existing cells. For insert, shift is Down or Right; for delete, Up or Left.',
    {
      range: z.string().describe('A1 range address to insert at or delete.'),
      action: z.enum(['insert', 'delete']).describe('Whether to insert or delete.'),
      shift: z
        .enum(['Down', 'Right', 'Up', 'Left'])
        .optional()
        .describe('Shift direction. Defaults: insert=Down, delete=Up.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
if (args.action === "insert") {
  range.insert(args.shift || "Down");
} else {
  range.delete(args.shift || "Up");
}
await context.sync();
return { action: args.action, range: args.range, shift: args.shift || (args.action === "insert" ? "Down" : "Up") };
`,
      ),
    ),
  )

  server.tool(
    'set_merge',
    'Merge or unmerge cells in a range.',
    {
      range: z.string().describe('A1 range address.'),
      action: z.enum(['merge', 'unmerge']).describe('Whether to merge or unmerge.'),
      across: z.boolean().optional().describe('merge: merge each row separately instead of one area. Default false.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
if (args.action === "merge") {
  range.merge(args.across === true);
} else {
  range.unmerge();
}
await context.sync();
return { action: args.action, range: args.range };
`,
      ),
    ),
  )

  server.tool(
    'autofit_range',
    'Autofit column widths and/or row heights of a range to their content.',
    {
      range: z.string().describe('A1 range address, e.g. "A:F".'),
      columns: z.boolean().optional().describe('Autofit column widths. Default true.'),
      rows: z.boolean().optional().describe('Autofit row heights. Default true.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var range = sheet.getRange(args.range);
if (args.columns !== false) range.format.autofitColumns();
if (args.rows !== false) range.format.autofitRows();
await context.sync();
return { autofit: args.range, columns: args.columns !== false, rows: args.rows !== false };
`,
      ),
    ),
  )

  server.tool(
    'add_comment',
    'Add a threaded comment to a cell.',
    {
      cellAddress: z.string().describe('Cell reference including sheet, e.g. "Sheet1!A1".'),
      text: z.string().describe('Comment text.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
context.workbook.comments.add(args.cellAddress, args.text);
await context.sync();
return { commented: args.cellAddress };
`,
      ),
    ),
  )

  server.tool(
    'refresh_pivottables',
    'Refresh all PivotTables in the workbook. Use after data changes when pivot summaries must be current.',
    {
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
context.workbook.pivotTables.refreshAll();
await context.sync();
return { refreshed: true };
`,
      ),
    ),
  )

  server.tool(
    'set_worksheet_protection',
    'Protect or unprotect a worksheet. Protection restricts user edits; cells formatted as locked are affected.',
    {
      sheetName: z.string().describe('Worksheet name.'),
      protect: z.boolean().describe('True to protect, false to unprotect.'),
      password: z.string().optional().describe('Optional password for the protection.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var sheet = context.workbook.worksheets.getItem(args.sheetName);
if (args.protect) {
  sheet.protection.protect(args.password);
} else {
  sheet.protection.unprotect(args.password);
}
await context.sync();
return { sheet: args.sheetName, protected: args.protect };
`,
      ),
    ),
  )

  server.tool(
    'read_threaded_comments',
    'Read all threaded comments in the workbook: content, author, creation date, resolved state, and replies. Use reply_to_comment or resolve_comment with a comment id from this output.',
    {
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var comments = context.workbook.comments;
comments.load("id,content,authorName,creationDate,resolved,replies");
await context.sync();
var replyLoads = [];
for (var i = 0; i < comments.items.length; i++) {
  comments.items[i].replies.load("id,content,authorName,creationDate");
  replyLoads.push(comments.items[i].replies);
}
await context.sync();
return {
  comments: comments.items.map(function (c) {
    return {
      id: c.id,
      content: c.content,
      author: c.authorName,
      created: c.creationDate,
      resolved: c.resolved,
      replies: c.replies.items.map(function (r) {
        return { id: r.id, content: r.content, author: r.authorName, created: r.creationDate };
      }),
    };
  }),
};
`,
      ),
    ),
  )

  server.tool(
    'reply_to_comment',
    'Add a reply to an existing threaded comment. Use the comment id from read_comments.',
    {
      commentId: z.string().describe('Comment id from read_comments.'),
      text: z.string().describe('Reply text.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var comment = context.workbook.comments.getItem(args.commentId);
comment.replies.add(args.text);
await context.sync();
return { replied: args.commentId };
`,
      ),
    ),
  )

  server.tool(
    'resolve_comment',
    'Resolve or reopen a threaded comment. Use the comment id from read_comments.',
    {
      commentId: z.string().describe('Comment id from read_comments.'),
      resolved: z.boolean().describe('True to resolve, false to reopen.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var comment = context.workbook.comments.getItem(args.commentId);
comment.resolved = args.resolved;
await context.sync();
comment.load("resolved");
await context.sync();
return { commentId: args.commentId, resolved: comment.resolved };
`,
      ),
    ),
  )

  server.tool(
    'set_hyperlink',
    'Set a hyperlink on a cell (address, display text, tooltip). The cell text is replaced by textToDisplay when provided. Pass address as null to remove the hyperlink.',
    {
      cellAddress: z.string().describe('Sheet-qualified cell reference, e.g. "Sheet1!B2".'),
      address: z
        .string()
        .nullable()
        .describe('Target URL, e.g. "https://example.com". Pass null to remove the hyperlink.'),
      textToDisplay: z.string().optional().describe('Cell text shown for the link. Defaults to the address.'),
      tooltip: z.string().optional().describe('Screen tip shown on hover.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var range = context.workbook.worksheets.getItem(args.cellAddress.split("!")[0]).getRange(args.cellAddress.split("!").slice(1).join("!"));
if (args.address === null) {
  range.hyperlink = null;
} else {
  var link = { address: args.address };
  if (args.textToDisplay !== undefined) link.textToDisplay = args.textToDisplay;
  if (args.tooltip !== undefined) link.screenTip = args.tooltip;
  range.hyperlink = link;
}
await context.sync();
range.load("hyperlink");
await context.sync();
return { cell: args.cellAddress, hyperlink: range.hyperlink === null ? null : { address: range.hyperlink.address, text: range.hyperlink.textToDisplay } };
`,
      ),
    ),
  )

  server.tool(
    'freeze_panes',
    'Freeze rows and columns of a worksheet at a cell (everything above and left of the cell stays visible), or unfreeze. E.g. freezeAt "A2" freezes the first row, "B2" freezes row 1 and column A.',
    {
      sheetName: z.string().describe('Worksheet name.'),
      freezeAt: z.string().nullable().describe('Cell reference to freeze at, e.g. "A2". Pass null to unfreeze.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
var sheet = context.workbook.worksheets.getItem(args.sheetName);
if (args.freezeAt === null) {
  sheet.freezePanes.unfreeze();
} else {
  sheet.freezePanes.freezeAt(sheet.getRange(args.freezeAt));
}
await context.sync();
return { sheet: args.sheetName, frozenAt: args.freezeAt };
`,
      ),
    ),
  )

  server.tool(
    'replace_in_sheet',
    'Find and replace text across a whole worksheet (substring match per cell by default). Returns the number of cells changed. Uses findAll so merged/hidden cells behave like the Excel UI.',
    {
      find: z.string().describe('Text to find.'),
      replace: z.string().describe('Replacement text.'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      matchCase: z.boolean().optional().describe('Case-sensitive search. Default false.'),
      completeMatch: z.boolean().optional().describe('Match the entire cell contents. Default false.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var live = [];
try {
  var found = sheet.findAll(args.find, { matchCase: args.matchCase === true, completeMatch: args.completeMatch === true });
  found.load("address,areas");
  await context.sync();
  found.areas.load("items/address,items/isNullObject,items/rowCount,items/columnCount,items/values");
  await context.sync();
  live = found.areas.items.filter(function (r) { return !r.isNullObject; });
} catch (e) {
  live = [];
}
var changed = 0;
for (var i = 0; i < live.length; i++) {
  var replaced = false;
  var grid = live[i].values;
  for (var r = 0; r < grid.length; r++) {
    for (var c = 0; c < grid[r].length; c++) {
      var v = grid[r][c];
      if (typeof v !== "string") continue;
      var hay = args.matchCase ? v : v.toLowerCase();
      var needle = args.matchCase ? args.find : args.find.toLowerCase();
      if (hay.indexOf(needle) === -1) continue;
      var out = "";
      var pos = 0;
      var idx = hay.indexOf(needle);
      while (idx !== -1) {
        out += v.slice(pos, idx) + args.replace;
        pos = idx + needle.length;
        idx = hay.indexOf(needle, pos);
      }
      out += v.slice(pos);
      grid[r][c] = out;
      replaced = true;
    }
  }
  if (replaced) {
    live[i].values = grid;
    changed++;
  }
}
await context.sync();
return { replacedCells: changed, matches: live.length };
`,
      ),
    ),
  )

  server.tool(
    'clear_range',
    'Clear the contents and formatting of a range in the live workbook.',
    {
      range: z.string().describe('A1 range address, e.g. "A1:C10".'),
      sheetName: z.string().optional().describe('Worksheet name. Defaults to the active worksheet.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
sheet.getRange(args.range).clear();
await context.sync();
return { cleared: args.range };
`,
      ),
    ),
  )

  server.tool(
    'copy_range',
    'Copy a source range to a target range in the live workbook. Copies values, formulas and formatting by default; valuesOnly copies just the results.',
    {
      source: z.string().describe('A1 range address to copy from, e.g. "A1:C3".'),
      target: z.string().describe('A1 range address (top-left corner) to copy into, e.g. "E1".'),
      sheetName: z.string().optional().describe('Worksheet name for both ranges. Defaults to the active worksheet.'),
      valuesOnly: z.boolean().optional().describe('Copy only the calculated values, not formulas or formatting.'),
      workbookId: z
        .string()
        .optional()
        .describe('Target workbook ID from list_workbooks. Optional when only one workbook is connected.'),
    },
    withTool(async (p) =>
      execExcel(
        pool,
        p.workbookId,
        p,
        `
${SHEET_LINE}
var src = sheet.getRange(args.source);
var dst = sheet.getRange(args.target);
if (args.valuesOnly) {
  dst.copyFrom(src, "Values");
} else {
  dst.copyFrom(src);
}
await context.sync();
return { copied: args.source + " -> " + args.target, valuesOnly: args.valuesOnly === true };
`,
      ),
    ),
  )
}
