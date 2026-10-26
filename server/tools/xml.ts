import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { DOMParser } from '@xmldom/xmldom'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import {
  applyZipEditAndReimport,
  escapeXml,
  exportSlide,
  extractSlideXmlFromZip,
  extractZipFiles,
  findShapeById,
  listZipPaths,
  NS_A,
  NS_P,
  parseSlideXml,
  reimportSlide,
  replaceShape,
  serializeXml,
  updateSlideXmlInZip,
} from '../xml-helpers.ts'
import { getConcurrentWarning } from './shared.ts'

export function registerXmlTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'edit_slide_xml',
    "Edit slide XML and reimport. Two modes: (1) xml mode — provide finished XML string (use read_slide_xml first to get current XML), (2) code mode — provide JS code that manipulates the pre-parsed DOM (receives: doc, findShapeById, NS_P, NS_A, escapeXml, serializeXml, DOMParser). Code mode preserves untouched attributes. The slide is exported, modified server-side, and reimported — data never enters Claude's context. NOTE: the reimport deletes and reinserts the slide, which makes PowerPoint's view jump to the first slide — reserve this for structural changes (new shapes, layout rebuilds).",
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based (slide 1 = first slide)'),
      xml: z
        .string()
        .optional()
        .describe(
          "Modified XML — full slide XML or a single shape's <p:sp> element (when shapeId is provided). Mutually exclusive with 'code'.",
        ),
      code: z
        .string()
        .optional()
        .describe(
          "JS code that manipulates the pre-parsed slide DOM. Receives: doc (Document), findShapeById(id) → Element|null, NS_P, NS_A (namespace strings), escapeXml(text), serializeXml(node), DOMParser. Mutually exclusive with 'xml'.",
        ),
      explanation: z
        .string()
        .optional()
        .describe('Brief description of what the code does (for logging, max 50 chars). Only used with code mode.'),
      shapeId: z
        .string()
        .optional()
        .describe(
          "Optional shape ID (xml mode only). If provided, replaces only that shape's <p:sp> element instead of the full slide XML.",
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, xml, code, shapeId, presentationId }) => {
      const slideIndex = slide - 1
      if ((!xml && !code) || (xml && code)) {
        return {
          content: [
            { type: 'text' as const, text: "Error: Provide either 'xml' or 'code', not both and not neither." },
          ],
          isError: true,
        }
      }

      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { zip, xmlString } = await extractSlideXmlFromZip(exported.base64)

      let finalXml: string
      let codeResultText: string | undefined
      if (code) {
        // Code mode: run agent-provided JS against the pre-parsed DOM
        const doc = parseSlideXml(xmlString)
        const sandbox = {
          doc,
          findShapeById: (id: string) => findShapeById(doc, id),
          NS_P,
          NS_A,
          escapeXml,
          serializeXml,
          DOMParser,
        }
        try {
          const keys = Object.keys(sandbox)
          const values = Object.values(sandbox)
          const fn = new Function(...keys, code)
          const codeResult = fn(...values)
          finalXml = serializeXml(doc)
          codeResultText = typeof codeResult === 'string' ? codeResult : undefined
        } catch (codeErr: unknown) {
          const msg = codeErr instanceof Error ? codeErr.message : String(codeErr)
          throw new Error(`Code execution error: ${msg}`)
        }
        finalXml = serializeXml(doc)
      } else if (shapeId) {
        const doc = parseSlideXml(xmlString)
        const shape = findShapeById(doc, shapeId)
        if (!shape) {
          throw new Error(`Shape with ID "${shapeId}" not found on slide ${slide}`)
        }
        replaceShape(doc, shape, xml!)
        finalXml = serializeXml(doc)
      } else {
        finalXml = xml!
      }

      // Pre-validate known reimport footguns so the agent gets an actionable
      // error instead of a bare InvalidArgument from Office.js.
      if (/overrideClrMapping/.test(finalXml)) {
        return {
          content: [
            {
              type: 'text' as const,
              text: 'Error: <a:overrideClrMapping> in p:clrMapOvr is rejected by the PowerPoint reimport (InvalidArgument). Use <a:masterClrMapping/> instead — theme colors resolve through the slide layout.',
            },
          ],
          isError: true,
        }
      }

      const modifiedBase64 = await updateSlideXmlInZip(zip, finalXml)
      await reimportSlide(
        pool,
        modifiedBase64,
        exported.slideId,
        exported.prevSlideId,
        target.ws,
        undefined,
        slideIndex,
      )
      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text =
        JSON.stringify({ success: true, ...(codeResultText ? { codeResult: codeResultText } : {}) }, null, 2) +
        (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'edit_slide_zip',
    'Update multiple files in the slide zip and reimport in a single operation. Accepts a map of { path: content } — can modify existing files or add new ones (e.g. chart XML + rels). Auto-registers Content_Types for new chart files. Use read_slide_zip first to get the current content.',
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based (slide 1 = first slide)'),
      files: z
        .record(z.string(), z.string())
        .describe(
          'Map of { zipPath: newContent }. Can include existing paths (to modify) or new paths (to add). Example: { "ppt/slides/slide1.xml": "<p:sld>...</p:sld>", "ppt/charts/chart1.xml": "<c:chartSpace>...</c:chartSpace>" }',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, files, presentationId }) => {
      const slideIndex = slide - 1
      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const newPaths = await applyZipEditAndReimport(pool, exported, files, target.ws, slideIndex)

      const warning = getConcurrentWarning(getSessionId(), target.presentationId, getActiveSessionCount())
      const text =
        JSON.stringify({ success: true, filesUpdated: Object.keys(files).length, newFiles: newPaths }, null, 2) +
        (warning ?? '')
      return { content: [{ type: 'text' as const, text }] }
    }),
  )

  server.tool(
    'read_slide_xml',
    "Read the full raw OOXML of a slide, or filter to a specific shape. Returns the slide's ppt/slides/slide1.xml content. Use with the /pptx skill's OOXML knowledge to understand the XML structure.",
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based (slide 1 = first slide), from scan_slide results'),
      shapeId: z
        .string()
        .optional()
        .describe("Optional shape ID to filter to. If provided, returns only that shape's <p:sp> element."),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, shapeId, presentationId }) => {
      const slideIndex = slide - 1
      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { xmlString } = await extractSlideXmlFromZip(exported.base64)

      if (shapeId) {
        const doc = parseSlideXml(xmlString)
        const shape = findShapeById(doc, shapeId)
        if (!shape) {
          throw new Error(`Shape with ID "${shapeId}" not found on slide ${slide}`)
        }
        return { content: [{ type: 'text' as const, text: serializeXml(shape) }] }
      }

      return { content: [{ type: 'text' as const, text: xmlString }] }
    }),
  )

  server.tool(
    'read_slide_zip',
    'Read multiple files from the exported slide zip. Returns slide XML, relationships, chart XMLs, and Content_Types. Use this to inspect chart data, rels, or other zip contents beyond what read_slide_xml provides. When no paths specified, auto-discovers all text/XML files in the zip.',
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based (slide 1 = first slide)'),
      paths: z
        .array(z.string())
        .optional()
        .describe(
          'Specific zip paths to read (e.g. ["ppt/slides/slide1.xml", "ppt/charts/chart1.xml"]). If omitted, auto-discovers all text/XML files.',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID from list_presentations. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, paths, presentationId }) => {
      const slideIndex = slide - 1
      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { zip, files } = await extractZipFiles(exported.base64, paths)
      const allPaths = listZipPaths(zip)
      const result = { zipContents: files, allPaths }
      return { content: [{ type: 'text' as const, text: JSON.stringify(result, null, 2) }] }
    }),
  )
}
