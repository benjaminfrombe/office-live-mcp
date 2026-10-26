import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { DOMParser, type Document, type Element } from '@xmldom/xmldom'
import { z } from 'zod'
import type { ConnectionPool } from '../bridge.ts'
import { withTool } from '../tool-helpers.ts'
import {
  escapeXml,
  exportSlide,
  extractSlideXmlFromZip,
  NS_A,
  NS_P,
  parseSlideXml,
  reimportSlide,
  serializeXml,
  updateSlideXmlInZip,
} from '../xml-helpers.ts'
import { getConcurrentWarning } from './shared.ts'

const PT = 12700 // EMU per point

// ---------------------------------------------------------------------------
// Geometry helpers — deterministic shape/connector math so agents never have
// to hand-compute OOXML bounding boxes, flips, or endpoint remapping.
// ---------------------------------------------------------------------------

interface Box {
  x: number // pt
  y: number
  w: number
  h: number
}

function getXfrm(el: Element): { xfrm: Element; off: Element; ext: Element } | null {
  // p:sp / p:pic / p:cxnSp carry <a:xfrm> inside p:spPr; p:graphicFrame carries <p:xfrm>
  let xfrm = el.getElementsByTagNameNS(NS_A, 'xfrm')[0]
  if (!xfrm) xfrm = el.getElementsByTagNameNS(NS_P, 'xfrm')[0]
  if (!xfrm) return null
  const off = xfrm.getElementsByTagNameNS(NS_A, 'off')[0]
  const ext = xfrm.getElementsByTagNameNS(NS_A, 'ext')[0]
  if (!off || !ext) return null
  return { xfrm, off, ext }
}

function readBox(el: Element): { box: Box; xfrm: Element; off: Element; ext: Element } | null {
  const g = getXfrm(el)
  if (!g) return null
  return {
    box: {
      x: Number.parseInt(g.off.getAttribute('x') ?? '0', 10) / PT,
      y: Number.parseInt(g.off.getAttribute('y') ?? '0', 10) / PT,
      w: Number.parseInt(g.ext.getAttribute('cx') ?? '0', 10) / PT,
      h: Number.parseInt(g.ext.getAttribute('cy') ?? '0', 10) / PT,
    },
    ...g,
  }
}

/** Shapes that behave like connectors: their vertical extent encodes endpoints. */
function isConnectorLike(el: Element): boolean {
  if (el.localName === 'cxnSp') return true
  const prst = el.getElementsByTagNameNS(NS_A, 'prstGeom')[0]?.getAttribute('prst') ?? ''
  if (/^(line|straightConnector\d+|bentConnector\d+|curvedConnector\d+|splineConnector\d*)$/.test(prst)) return true
  // A custGeom that is a single open stroke (one moveTo + one bezier/line) is a drawn connector.
  const cust = el.getElementsByTagNameNS(NS_A, 'custGeom')[0]
  if (cust) {
    const paths = cust.getElementsByTagNameNS(NS_A, 'path')
    for (const p of paths) {
      const moveTos = p.getElementsByTagNameNS(NS_A, 'moveTo')
      const curves = p.getElementsByTagNameNS(NS_A, 'cubicBezTo')
      const lines = p.getElementsByTagNameNS(NS_A, 'lnTo')
      if (moveTos.length === 1 && curves.length + lines.length === 1) return true
    }
  }
  return false
}

/** Trim a center-to-center ray to the rectangle boundary of a shape. */
function edgePoint(box: Box, toward: { x: number; y: number }): { x: number; y: number } {
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const dx = toward.x - cx
  const dy = toward.y - cy
  if (dx === 0 && dy === 0) return { x: cx, y: cy }
  const tx = dx !== 0 ? box.w / 2 / Math.abs(dx) : Number.POSITIVE_INFINITY
  const ty = dy !== 0 ? box.h / 2 / Math.abs(dy) : Number.POSITIVE_INFINITY
  const t = Math.min(tx, ty)
  return { x: cx + dx * t, y: cy + dy * t }
}

function nextShapeId(doc: Document): number {
  let max = 1
  for (const c of doc.getElementsByTagNameNS(NS_P, 'cNvPr')) {
    const v = Number.parseInt(c.getAttribute('id') ?? '1', 10)
    if (Number.isFinite(v) && v > max) max = v
  }
  return max + 1
}

function connectorXml(
  id: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  opts: { color: string; weightPt: number; alphaPct?: number; curve: 'straight' | 'bezier' },
): string {
  const X = Math.min(x1, x2)
  const Y = Math.min(y1, y2)
  const W = Math.abs(x2 - x1)
  const H = Math.abs(y2 - y1)
  const alpha = opts.alphaPct !== undefined ? `<a:alpha val="${Math.round(opts.alphaPct * 1000)}"/>` : ''
  const ln = `<a:ln w="${Math.round(opts.weightPt * PT)}" cap="rnd"><a:solidFill><a:srgbClr val="${opts.color}">${alpha}</a:srgbClr></a:solidFill></a:ln>`
  let geom: string
  if (opts.curve === 'bezier') {
    // Horizontal tangents at both ends; path encodes direction explicitly, so
    // NO flip attributes (flipV on a custGeom double-mirrors the path).
    const c1x = x1 + (x2 - x1) * 0.4
    const c2x = x2 - (x2 - x1) * 0.4
    const pt = (x: number, y: number) => `<a:pt x="${Math.round((x - X) * PT)}" y="${Math.round((y - Y) * PT)}"/>`
    geom =
      `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:rect l="l" t="t" r="r" b="b"/>` +
      `<a:pathLst><a:path w="${Math.round(W * PT)}" h="${Math.round(Math.max(H, 1 / PT) * PT)}" fill="none">` +
      `<a:moveTo>${pt(x1, y1)}</a:moveTo>` +
      `<a:cubicBezTo>${pt(c1x, y1)}${pt(c2x, y2)}${pt(x2, y2)}</a:cubicBezTo>` +
      `</a:path></a:pathLst></a:custGeom>`
  } else {
    // prst="line" draws TL->BR of its box; flips select the actual start corner.
    const flipH = x1 > x2 ? ' flipH="1"' : ''
    const flipV = y1 > y2 ? ' flipV="1"' : ''
    geom = `<a:prstGeom prst="line"${flipH}${flipV}><a:avLst/></a:prstGeom>`
  }
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(X * PT)}" y="${Math.round(Y * PT)}"/>` +
    `<a:ext cx="${Math.round(W * PT)}" cy="${Math.round(H * PT)}"/></a:xfrm>${geom}${ln}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="nl-NL"/></a:p></p:txBody></p:sp>`
  )
}

const endpointSchema = z
  .union([
    z.object({
      shapeId: z.string().describe('Attach to this shape (line meets its bounding-box edge facing the other end).'),
    }),
    z.object({ x: z.number(), y: z.number() }).describe('Absolute point in points (slide is 960x540 pt).'),
  ])
  .describe('Endpoint: {shapeId} or {x, y}')

export function registerGeometryTools(
  server: McpServer,
  pool: ConnectionPool,
  getSessionId: () => string | undefined,
  getActiveSessionCount: () => number,
): void {
  server.tool(
    'add_connector',
    'Draw a line or curved connector between two shapes or points, with all OOXML bounding-box math (off/ext/flips, bezier path) computed for you. For bezier curves the path leaves and arrives horizontally — no flip attributes are used, avoiding the custGeom double-mirror pitfall. The shape is appended on top of the slide.',
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based'),
      from: endpointSchema,
      to: endpointSchema,
      color: z.string().optional().describe('Line color as RRGGBB hex, e.g. "EF463C". Default EF463C.'),
      weightPt: z.number().optional().describe('Line width in points. Default 1.'),
      alphaPct: z.number().min(0).max(100).optional().describe('Opacity in percent (0-100). Omit for fully opaque.'),
      curve: z
        .enum(['straight', 'bezier'])
        .optional()
        .describe('straight (default) or bezier (horizontal tangents at both ends).'),
      zOrder: z
        .enum(['top', 'bottom'])
        .optional()
        .describe('"top" (default) appends above all shapes; "bottom" inserts below all shapes.'),
      beforeShapeId: z
        .string()
        .optional()
        .describe('Insert the connector directly below this shape ID (z-order). Overrides zOrder.'),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, from, to, color, weightPt, alphaPct, curve, zOrder, beforeShapeId, presentationId }) => {
      const slideIndex = slide - 1
      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { zip, xmlString } = await extractSlideXmlFromZip(exported.base64)
      const doc = parseSlideXml(xmlString)

      const resolveEnd = async (end: { shapeId?: string; x?: number; y?: number }) => {
        if (typeof end.x === 'number' && typeof end.y === 'number') return { x: end.x, y: end.y }
        const id = (end as { shapeId?: string }).shapeId
        if (!id) throw new Error('Endpoint must have either shapeId or x/y')
        for (const tag of ['sp', 'pic', 'graphicFrame', 'cxnSp']) {
          for (const el of doc.getElementsByTagNameNS(NS_P, tag)) {
            const cnv = el.getElementsByTagNameNS(NS_P, 'cNvPr')[0]
            if (cnv?.getAttribute('id') === id) {
              const rb = readBox(el)
              if (!rb) throw new Error(`Shape "${id}" has no geometry`)
              return rb.box
            }
          }
        }
        throw new Error(`Shape with ID "${id}" not found on slide ${slide}`)
      }

      const fromBox = await resolveEnd(from)
      const toBox = await resolveEnd(to)
      const fromIsBox = !('x' in from) && !('y' in from)
      const toIsBox = !('x' in to) && !('y' in to)
      const toCenter =
        'w' in toBox ? { x: toBox.x + toBox.w / 2, y: toBox.y + toBox.h / 2 } : (toBox as { x: number; y: number })
      const fromCenter =
        'w' in fromBox
          ? { x: fromBox.x + fromBox.w / 2, y: fromBox.y + fromBox.h / 2 }
          : (fromBox as { x: number; y: number })
      const p1 = fromIsBox ? edgePoint(fromBox as Box, toCenter) : (fromBox as { x: number; y: number })
      const p2 = toIsBox ? edgePoint(toBox as Box, fromCenter) : (toBox as { x: number; y: number })

      const createdId = nextShapeId(doc)
      const xml = connectorXml(createdId, p1.x, p1.y, p2.x, p2.y, {
        color: (color ?? 'EF463C').replace('#', '').toUpperCase(),
        weightPt: weightPt ?? 1,
        alphaPct,
        curve: curve ?? 'straight',
      })
      const wrapper = new DOMParser().parseFromString(
        `<root xmlns:p="${NS_P}" xmlns:a="${NS_A}">${xml}</root>`,
        'text/xml',
      )
      const imported = wrapper.documentElement?.firstChild as Element | undefined
      if (!imported) throw new Error('Failed to build connector XML')
      const spTree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0]
      const node = doc.importNode(imported, true)
      let position = 'top'
      if (beforeShapeId) {
        let anchor: Element | null = null
        for (const tag of ['sp', 'pic', 'graphicFrame', 'cxnSp']) {
          for (const el of spTree.getElementsByTagNameNS(NS_P, tag)) {
            const cnv = el.getElementsByTagNameNS(NS_P, 'cNvPr')[0]
            if (cnv?.getAttribute('id') === beforeShapeId && el.parentNode === spTree) {
              anchor = el
              break
            }
          }
          if (anchor) break
        }
        if (!anchor) throw new Error(`beforeShapeId "${beforeShapeId}" not found on slide ${slide}`)
        spTree.insertBefore(node, anchor)
        position = `below shape ${beforeShapeId}`
      } else if (zOrder === 'bottom') {
        spTree.insertBefore(node, spTree.firstChild)
        position = 'bottom'
      } else {
        spTree.appendChild(node)
      }

      const modifiedBase64 = await updateSlideXmlInZip(zip, serializeXml(doc))
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
      return {
        content: [
          {
            type: 'text' as const,
            text:
              JSON.stringify(
                {
                  success: true,
                  connectorId: String(createdId),
                  position,
                  from: p1,
                  to: p2,
                  curve: curve ?? 'straight',
                },
                null,
                2,
              ) + (warning ?? ''),
          },
        ],
      }
    }),
  )

  server.tool(
    'transform_slide',
    'Vertically rescale a band of shapes (e.g. shrink the content of a slide by 30%) in ONE atomic operation. Shapes whose top falls inside [topPt, bottomPt] are remapped to start at newTopPt with factor scaleY. Connector-like shapes (cxnSp, line/connector geometries, single-stroke bezier custGeom) get their height rescaled too so endpoints stay attached; all other shapes keep their height unless listed in scaleHeightShapeIds. Text, images and footers outside the band are untouched.',
    {
      slide: z.number().int().min(1).describe('Slide number, 1-based'),
      topPt: z.number().describe('Top of the band to transform, in points (slide is 960x540 pt).'),
      bottomPt: z.number().describe('Bottom of the band to transform, in points.'),
      scaleY: z.number().describe('Vertical scale factor, e.g. 0.7 = 30% shorter.'),
      newTopPt: z.number().optional().describe('Where the top of the scaled band should land. Default: keep topPt.'),
      scaleHeightShapeIds: z
        .array(z.string())
        .optional()
        .describe(
          'Shape IDs (from scan_slide) whose height should scale with the band (e.g. the two big panel rectangles).',
        ),
      presentationId: z
        .string()
        .optional()
        .describe('Target presentation ID. Optional when only one presentation is connected.'),
    },
    withTool(async ({ slide, topPt, bottomPt, scaleY, newTopPt, scaleHeightShapeIds, presentationId }) => {
      const slideIndex = slide - 1
      const target = pool.resolveTarget(presentationId)
      const exported = await exportSlide(pool, slideIndex, target.ws)
      const { zip, xmlString } = await extractSlideXmlFromZip(exported.base64)
      const doc = parseSlideXml(xmlString)

      const anchor = newTopPt ?? topPt
      const eps = 0.01
      const scaleIds = new Set(scaleHeightShapeIds ?? [])
      const spTree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0]
      const tags = ['sp', 'pic', 'graphicFrame', 'cxnSp']

      // Pass 1: map positions (and heights where applicable), recording old/new
      // boxes so pass 2 can re-anchor connector endpoints to shape centers.
      const geom = new Map<string, { old: Box; box: Box; isConn: boolean; el: Element }>()
      for (const tag of tags) {
        for (const el of [...spTree.getElementsByTagNameNS(NS_P, tag)]) {
          if (el.parentNode !== spTree) continue
          const rb = readBox(el)
          if (!rb) continue
          if (rb.box.y < topPt - eps || rb.box.y > bottomPt + eps) continue
          const cnv = el.getElementsByTagNameNS(NS_P, 'cNvPr')[0]
          const sid = cnv?.getAttribute('id') ?? '?'
          const oldBox = { ...rb.box }
          rb.off.setAttribute('y', String(Math.round((anchor + (rb.box.y - topPt) * scaleY) * PT)))
          const isConn = isConnectorLike(el)
          if (isConn) {
            rb.ext.setAttribute('cy', String(Math.round(rb.box.h * scaleY * PT)))
          } else if (scaleIds.has(sid)) {
            rb.ext.setAttribute('cy', String(Math.round(rb.box.h * scaleY * PT)))
          }
          // Explicit NEW box (rb.box was captured before the writes above and
          // still holds old geometry — using it as "new" made snapping write
          // old positions back over pass 1's connector updates).
          const newBox = {
            x: rb.box.x,
            y: anchor + (rb.box.y - topPt) * scaleY,
            w: rb.box.w,
            h: isConn || scaleIds.has(sid) ? rb.box.h * scaleY : rb.box.h,
          }
          geom.set(sid, { old: oldBox, box: newBox, isConn, el })
        }
      }

      // Center lookup for endpoint snapping: old center -> new center.
      const centers: Array<{ ol: number; or: number; oy: number; ny: number }> = []
      for (const g of geom.values()) {
        if (g.isConn) continue
        centers.push({
          ol: g.old.x,
          or: g.old.x + g.old.w,
          oy: g.old.y + g.old.h / 2,
          ny: g.box.y + g.box.h / 2,
        })
      }
      // Endpoints attach at shape EDGES, not centers: match by horizontal
      // overlap with the shape's span, then nearest old center vertically.
      const snapY = (oy: number, ox: number): number | null => {
        let best: number | null = null
        let bestD = 2 // pt tolerance
        for (const c of centers) {
          if (ox < c.ol - 2 || ox > c.or + 2) continue
          const d = Math.abs(c.oy - oy)
          if (d < bestD) {
            bestD = d
            best = c.ny
          }
        }
        return best
      }
      const mapY = (oy: number): number => anchor + (oy - topPt) * scaleY

      // Pass 2: re-anchor connector endpoints. Endpoint drift is otherwise
      // halfheight * (scaleY - 1), because pill centers are top + h/2 and
      // heights do not scale.
      const connectorsRescaled: string[] = []
      for (const g of geom.values()) {
        if (!g.isConn) continue
        connectorsRescaled.push(g.el.getElementsByTagNameNS(NS_P, 'cNvPr')[0]?.getAttribute('id') ?? '?')
        const { old, box } = g
        const prst = g.el.getElementsByTagNameNS(NS_A, 'prstGeom')[0]?.getAttribute('prst') ?? ''
        let e1: { x: number; y: number }
        let e2: { x: number; y: number }
        if (prst === 'line' || g.el.localName === 'cxnSp') {
          const xf = g.el.getElementsByTagNameNS(NS_A, 'xfrm')[0]
          const flipH = xf.getAttribute('flipH') === '1'
          const flipV = xf.getAttribute('flipV') === '1'
          e1 = { x: flipH ? old.x + old.w : old.x, y: flipV ? old.y + old.h : old.y }
          e2 = { x: flipH ? old.x : old.x + old.w, y: flipV ? old.y : old.y + old.h }
        } else {
          // custGeom single stroke: endpoints = first moveTo and last path point,
          // mapped from path units to slide coordinates.
          const cust = g.el.getElementsByTagNameNS(NS_A, 'custGeom')[0]
          const path = cust?.getElementsByTagNameNS(NS_A, 'path')[0]
          if (!path) continue
          const pw = Number.parseInt(path.getAttribute('w') ?? '1', 10) || 1
          const ph = Number.parseInt(path.getAttribute('h') ?? '1', 10) || 1
          const pts = path.getElementsByTagNameNS(NS_A, 'pt')
          const first = pts[0]
          const last = pts[pts.length - 1]
          if (!first || !last) continue
          e1 = {
            x: old.x + (Number.parseInt(first.getAttribute('x') ?? '0', 10) / pw) * old.w,
            y: old.y + (Number.parseInt(first.getAttribute('y') ?? '0', 10) / ph) * old.h,
          }
          e2 = {
            x: old.x + (Number.parseInt(last.getAttribute('x') ?? '0', 10) / pw) * old.w,
            y: old.y + (Number.parseInt(last.getAttribute('y') ?? '0', 10) / ph) * old.h,
          }
        }
        const n1 = { x: e1.x, y: snapY(e1.y, e1.x) ?? mapY(e1.y) }
        const n2 = { x: e2.x, y: snapY(e2.y, e2.x) ?? mapY(e2.y) }
        const X = Math.min(n1.x, n2.x)
        const Y = Math.min(n1.y, n2.y)
        const W = Math.abs(n2.x - n1.x)
        const H = Math.abs(n2.y - n1.y)
        const rb = readBox(g.el)
        if (!rb) continue
        rb.off.setAttribute('x', String(Math.round(X * PT)))
        rb.off.setAttribute('y', String(Math.round(Y * PT)))
        rb.ext.setAttribute('cx', String(Math.round(W * PT)))
        rb.ext.setAttribute('cy', String(Math.round(H * PT)))
        if (prst === 'line' || g.el.localName === 'cxnSp') {
          const xf = g.el.getElementsByTagNameNS(NS_A, 'xfrm')[0]
          if (n2.x < n1.x) xf.setAttribute('flipH', '1')
          else xf.removeAttribute('flipH')
          if (n2.y < n1.y) xf.setAttribute('flipV', '1')
          else xf.removeAttribute('flipV')
        } else {
          // custGeom: map every path point with the same affine y-map, then
          // override the two endpoints with the snapped centers, and resync
          // path units to the new box.
          const cust = g.el.getElementsByTagNameNS(NS_A, 'custGeom')[0]
          const path = cust?.getElementsByTagNameNS(NS_A, 'path')[0]
          if (!path) continue
          const pw = Number.parseInt(path.getAttribute('w') ?? '1', 10) || 1
          const ph = Number.parseInt(path.getAttribute('h') ?? '1', 10) || 1
          const pts = path.getElementsByTagNameNS(NS_A, 'pt')
          for (let i = 0; i < pts.length; i++) {
            const pt = pts[i]
            const px = Number.parseInt(pt.getAttribute('x') ?? '0', 10)
            const py = Number.parseInt(pt.getAttribute('y') ?? '0', 10)
            const ax = old.x + (px / pw) * old.w
            const ay = old.y + (py / ph) * old.h
            let ny = mapY(ay)
            if (i === 0) ny = n1.y
            if (i === pts.length - 1) ny = n2.y
            pt.setAttribute('x', String(Math.round((ax - X) * PT)))
            pt.setAttribute('y', String(Math.round((ny - Y) * PT)))
          }
          path.setAttribute('w', String(Math.round(W * PT)))
          path.setAttribute('h', String(Math.round(H * PT)))
        }
      }

      const heightScaled: string[] = []
      for (const [sid, g] of geom.entries()) {
        if (!g.isConn && scaleIds.has(sid)) heightScaled.push(sid)
      }

      const modifiedBase64 = await updateSlideXmlInZip(zip, serializeXml(doc))
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
      return {
        content: [
          {
            type: 'text' as const,
            text:
              JSON.stringify({ success: true, moved: geom.size, heightScaled, connectorsRescaled }, null, 2) +
              (warning ?? ''),
          },
        ],
      }
    }),
  )
}

// escapeXml is imported for parity with sibling modules; keep tree-shaking honest.
void escapeXml
