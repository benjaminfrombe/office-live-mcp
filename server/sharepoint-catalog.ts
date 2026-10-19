import { readFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { join, resolve } from 'node:path'
import { substituteManifestPort } from './manifest.ts'

/**
 * Directory holding the manifest XML files (the bridge's addin/ dir), set by
 * the server at startup so file-content endpoints can serve the manifests.
 */
let manifestDir = ''
export function setManifestDir(dir: string): void {
  manifestDir = dir
}

function readManifest(req: IncomingMessage, serverRelativeUrl: string): string | null {
  const name = serverRelativeUrl.split('/').pop() ?? ''
  if (!manifestDir) return null
  try {
    const raw = readFileSync(join(resolve(manifestDir), name), 'utf-8')
    const host = String(req.headers.host ?? '')
    const portMatch = /:(\d+)$/.exec(host)
    const targetPort = portMatch ? Number.parseInt(portMatch[1], 10) : 8443
    return substituteManifestPort(raw, 8080, targetPort)
  } catch {
    return null
  }
}

/**
 * Minimal SharePoint app-catalog emulation.
 *
 * Recent Office builds only accept HTTPS URL catalogs and treat them as
 * SharePoint sites: Office enumerates the catalog via the SharePoint REST
 * API (`{catalogUrl}/_api/web/lists`, then the catalog list items) instead
 * of reading a plain manifest file. This module answers those requests so a
 * plain bridge can act as a self-hosted app catalog: the manifests it
 * already serves become available in every document, the way the wef
 * folder makes them available on macOS.
 *
 * The protocol is reverse-engineered from Office's requests; every `_api`
 * call is logged with its Accept header so unexpected follow-ups can be
 * implemented one by one.
 */

interface CatalogApp {
  /** Path served by the bridge's static handler, with port substitution. */
  serverRelativeUrl: string
  title: string
  id: string
}

// Office add-in manifests live in the "Apps for Office" document library of
// an App Catalog site (the "Apps for SharePoint" library holds .app packages,
// which is NOT what the Office client enumerates for web add-ins).
const CATALOG_LIST_TITLE = 'Apps for Office'
const CATALOG_LIST_ID = '9a1c5e52-6d42-4b0a-9f47-3d2f1c8b7a01'

// Properties that SharePoint models as navigation properties: in Atom feeds
// they are emitted as related links, never inside m:properties. File is the
// exception: Office reads File.ServerRelativeUrl from the item properties,
// so File must stay a complex property (atomProperty emits it inline).
const NAVIGATION_KEYS = new Set(['ParentList', 'FirstUniqueAncestorSecurableObject', 'RootFolder'])

const CATALOG_APPS: CatalogApp[] = [
  {
    // VM minimal variant: no VersionOverrides, to isolate catalog acceptance
    // from manifest schema strictness on Windows Office builds.
    serverRelativeUrl: '/manifest.vm-minimal.xml',
    title: 'Office Live MCP Word',
    id: '6B03BAEA-D5F8-43A1-B136-791D9D1AB231',
  },
  {
    serverRelativeUrl: '/manifest.vm-excel.xml',
    title: 'Office Live MCP Excel',
    id: 'B116267F-F3C1-46D4-9A06-1FE8D7F13A69',
  },
  {
    serverRelativeUrl: '/manifest.vm-powerpoint.xml',
    title: 'Office Live MCP PowerPoint',
    id: 'AE89909C-2813-4B08-9E1B-49E7379BD0E6',
  },
]

/**
 * SharePoint REST defaults to Atom XML unless the client explicitly asks
 * for JSON. Office's catalog enumerator sends requests without an Accept
 * header and expects Atom; answering those with JSON makes it retry and
 * eventually report "Cannot connect to catalog".
 */
function wantsAtom(req: IncomingMessage): boolean {
  const accept = String(req.headers.accept ?? '')
  return !accept.includes('json')
}

function sendJson(res: ServerResponse, body: unknown): void {
  res.writeHead(200, { 'Content-Type': 'application/json;odata=nometadata;charset=utf-8' })
  res.end(JSON.stringify(body))
}

function atomEscape(value: unknown): string {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

function atomProperty(key: string, value: unknown): string {
  if (value === null || value === undefined) {
    return `<d:${key} m:null="true" />`
  }
  if (typeof value === 'boolean') {
    return `<d:${key} m:type="Edm.Boolean">${value}</d:${key}>`
  }
  if (typeof value === 'number' && Number.isInteger(value)) {
    return `<d:${key} m:type="Edm.Int32">${value}</d:${key}>`
  }
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    const record = value as Record<string, unknown>
    // SP.FieldUrlValue: <d:Prop m:type="SP.FieldUrlValue"><d:Description/><d:Url/></d:Prop>
    if ('Description' in record && 'Url' in record) {
      const desc = atomEscape(record.Description ?? '')
      const url = atomEscape(record.Url ?? '')
      return `<d:${key} m:type="SP.FieldUrlValue"><d:Description>${desc}</d:Description><d:Url>${url}</d:Url></d:${key}>`
    }
    // Other complex values: wrapped nested properties. This is the form
    // Office's catalog parser actually reads File.ServerRelativeUrl from
    // (empirically verified: with this wrapping Office downloads the item's
    // manifest file; with the "cleaner" inline complex form it does not).
    const inner = Object.entries(record)
      .map(([k, v]) => atomProperty(k, v))
      .join('')
    return `<d:${key}><m:properties>${inner}</m:properties></d:${key}>`
  }
  const strValue = atomEscape(value)
  const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(strValue)
  if (isGuid) {
    return `<d:${key} m:type="Edm.Guid">${strValue}</d:${key}>`
  }
  return `<d:${key}>${strValue}</d:${key}>`
}

function sendAtom(res: ServerResponse, req: IncomingMessage, entries: Array<Record<string, unknown>>): void {
  // SharePoint-style Atom feed: Office's catalog enumerator parses this when
  // it sends requests without a JSON Accept header.
  const base = `https://${req.headers.host ?? 'localhost:8443'}`
  const now = new Date().toISOString()
  const entriesXml = entries
    .map((entry, i) => {
      const props = Object.entries(entry)
        // Navigation properties (File, ParentList, ...) are not data
        // properties in SharePoint Atom; they are emitted as related links.
        .filter(([k]) => !NAVIGATION_KEYS.has(k))
        .map(([k, v]) => atomProperty(k, v))
        .join('')
      const title = atomEscape(entry.Title ?? `item-${i}`)
      const isListItem = 'GUID' in entry
      const editPath = isListItem
        ? `Web/Lists(guid'${CATALOG_LIST_ID}')/Items(${atomEscape(entry.Id ?? i)})`
        : `Web/Lists(guid'${atomEscape(entry.Id ?? i)}')`
      const term = isListItem ? 'SP.Data.Apps_x0020_for_x0020_OfficeListItem' : 'SP.List'
      const fileUrl = typeof entry.FileRef === 'string' ? entry.FileRef : null
      const fileLink =
        isListItem && fileUrl
          ? `<link rel="http://schemas.microsoft.com/ado/2007/08/dataservices/related/File" type="application/atom+xml;type=entry" title="File" href="${editPath}/File" />`
          : ''
      const parentListLink = isListItem
        ? `<link rel="http://schemas.microsoft.com/ado/2007/08/dataservices/related/ParentList" type="application/atom+xml;type=entry" title="ParentList" href="Web/Lists(guid'${CATALOG_LIST_ID}')" />`
        : ''
      return (
        `<entry>` +
        `<id>${base}/_api/${editPath}</id>` +
        `<category term="${term}" scheme="http://schemas.microsoft.com/ado/2007/08/dataservices/scheme" />` +
        `<link rel="edit" href="${editPath}" />` +
        fileLink +
        parentListLink +
        `<title type="text">${title}</title>` +
        `<updated>${now}</updated>` +
        `<author><name>Office Live MCP Bridge</name></author>` +
        `<content type="application/xml"><m:properties>${props}</m:properties></content>` +
        `</entry>`
      )
    })
    .join('')
  const itemsFeed = entries.length > 0 && 'GUID' in entries[0]
  const feedPath = itemsFeed ? `Web/Lists(guid'${CATALOG_LIST_ID}')/Items` : 'Web/Lists'
  res.writeHead(200, { 'Content-Type': 'application/atom+xml;charset=utf-8' })
  res.end(
    `<?xml version="1.0" encoding="utf-8"?>` +
      `<feed xml:base="${base}/_api/" xmlns="http://www.w3.org/2005/Atom" ` +
      `xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" ` +
      `xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">` +
      `<id>${base}/_api/${feedPath}</id>` +
      `<title type="text">${itemsFeed ? 'Items' : 'Lists'}</title>` +
      `<updated>${now}</updated>` +
      `<link rel="self" href="${feedPath}" />` +
      entriesXml +
      `</feed>`,
  )
}

function sendCollection(res: ServerResponse, req: IncomingMessage, value: Array<Record<string, unknown>>): void {
  if (wantsAtom(req)) {
    sendAtom(res, req, value)
    return
  }
  sendJson(res, { value })
}

function listItem(app: CatalogApp, index: number): Record<string, unknown> {
  return {
    FileSystemObjectType: 0,
    Id: index + 1,
    GUID: app.id,
    Title: app.title,
    FileLeafRef: app.serverRelativeUrl.split('/').pop(),
    FileRef: app.serverRelativeUrl,
    ContentTypeId: '0x01010C009A2929A6E2164D40AC668301D8D91CE5',
    AppProductID: app.id,
    AppVersion: '1.0.0.0',
    AppPublisher: 'Office Live MCP Bridge',
    AppDescription: 'Live Office.js bridge add-in',
    AppShortDescription: 'Live Office.js bridge add-in',
    AppThumbnailURL: { Description: '', Url: '' },
    AppSupportURL: null,
    AppVideoURL: null,
    IsAppPackageEnabled: true,
    IsValidAppPackage: true,
    AppPackageErrorMessage: 'No errors.',
    SharePointAppCategory: null,
    IsFeaturedApp: false,
    AppMetadataLocale: '1033',
    IsDefaultAppMetadataLocale: true,
    ServerRedirectedEmbedUri: null,
    ServerRedirectedEmbedUrl: '',
    File: {
      ServerRelativeUrl: app.serverRelativeUrl,
      Name: app.serverRelativeUrl.split('/').pop(),
      Title: app.title,
      Length: '4096',
      TimeLastModified: '2026-10-19T18:00:00Z',
    },
    FirstUniqueAncestorSecurableObject: 1,
    ParentList: { Id: CATALOG_LIST_ID, Title: CATALOG_LIST_TITLE },
  }
}

/**
 * Handle a SharePoint `_api` request. Returns true when the request was
 * handled (the caller should not run other handlers).
 */
export function handleSharePointCatalog(req: IncomingMessage, res: ServerResponse, rawUrl: string): boolean {
  const method = (req.method ?? 'GET').toUpperCase()
  const qIndex = rawUrl.indexOf('?')
  const path = qIndex === -1 ? rawUrl : rawUrl.slice(0, qIndex)
  const query = qIndex === -1 ? '' : rawUrl.slice(qIndex)

  const apiIndex = path.indexOf('/_api/')
  if (apiIndex === -1) return false

  const apiPath = path.slice(apiIndex + '/_api/'.length)
  console.error(`[catalog] ${method} ${path}${query} accept=${req.headers.accept ?? '(none)'}`)

  // Normalize: strip a trailing slash and the `Web` alias used by some clients.
  const normalized = apiPath.replace(/\/+$/, '')
  const lower = normalized.toLowerCase()

  // Site/web root — Office sometimes probes the web first.
  if (lower === 'web' || lower === 'web/lists' || lower === 'web/alllists') {
    if (lower === 'web') {
      sendCollection(res, req, [{ Id: '11111111-2222-3333-4444-555555555555', Title: 'o365 MCP Catalog', Url: '/' }])
      return true
    }
    sendCollection(res, req, [
      {
        Id: CATALOG_LIST_ID,
        Title: CATALOG_LIST_TITLE,
        BaseTemplate: 101,
        ItemCount: CATALOG_APPS.length,
        Description: 'o365 MCP self-hosted add-in catalog',
        ParentWebUrl: '/',
        Hidden: false,
        EnableAttachments: false,
        ListItemEntityTypeFullName: 'SP.Data.Apps_x0020_for_x0020_OfficeListItem',
        RootFolder: { ServerRelativeUrl: '/Apps for Office' },
      },
    ])
    return true
  }

  // List items — by title, by guid with slashes, or guid in parentheses
  // (`web/lists(guid'...')/items` is what Office's enumerator actually sends).
  const itemsMatch =
    /^web\/lists\/(?:getbytitle\('([^']+)'\)|guid'([^']+)')\/items$/.exec(normalized) ??
    /^web\/lists\(guid'([^']+)'\)\/items$/.exec(normalized)
  if (itemsMatch) {
    const items = CATALOG_APPS.map((app, i) => listItem(app, i))
    sendCollection(res, req, items)
    return true
  }

  // The list itself.
  const listMatch =
    /^web\/lists\/(?:getbytitle\('([^']+)'\)|guid'([^']+)'|bytitle\('([^']+)'\))$/.exec(normalized) ??
    /^web\/lists\(guid'([^']+)'\)$/.exec(normalized)
  if (listMatch) {
    sendCollection(res, req, [
      {
        Id: CATALOG_LIST_ID,
        Title: CATALOG_LIST_TITLE,
        BaseTemplate: 101,
        ItemCount: CATALOG_APPS.length,
        Description: 'o365 MCP self-hosted add-in catalog',
        ParentWebUrl: '/',
        Hidden: false,
        EnableAttachments: false,
        ListItemEntityTypeFullName: 'SP.Data.Apps_x0020_for_x0020_OfficeListItem',
        RootFolder: { ServerRelativeUrl: '/Apps for Office' },
      },
    ])
    return true
  }

  // File content — Office downloads the manifest either through the item's
  // File navigation property or via getfilebyserverrelativeurl.
  const itemFileMatch = /^web\/lists\(guid'([^']+)'\)\/items\((\d+)\)\/file(?:\/\$value)?$/.exec(normalized)
  if (itemFileMatch) {
    const app = CATALOG_APPS[Number.parseInt(itemFileMatch[2], 10) - 1]
    if (app && normalized.endsWith('/$value')) {
      const content = readManifest(req, app.serverRelativeUrl)
      if (content !== null) {
        res.writeHead(200, { 'Content-Type': 'application/xml;charset=utf-8' })
        res.end(content)
        return true
      }
    }
    if (app) {
      sendCollection(res, req, [
        listItem(app, Number.parseInt(itemFileMatch[2], 10) - 1).File as Record<string, unknown>,
      ])
      return true
    }
  }

  const getFileMatch = /^web\/getfilebyserverrelativeurl\((.+)\)(?:\/\$value)?$/.exec(normalized)
  if (getFileMatch) {
    let url = getFileMatch[1].replace(/^['"]|['"]$/g, '')
    try {
      url = decodeURIComponent(url)
    } catch {
      // keep raw
    }
    const app = CATALOG_APPS.find((a) => a.serverRelativeUrl === url)
    const content = app ? readManifest(req, app.serverRelativeUrl) : null
    if (content !== null) {
      res.writeHead(200, { 'Content-Type': 'application/xml;charset=utf-8' })
      res.end(content)
      return true
    }
    console.error(`[catalog] getfilebyserverrelativeurl: unknown path ${url}`)
    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('404 Not Found')
    return true
  }

  // Unknown follow-up: log clearly and answer with an empty success so the
  // conversation continues and the next request reveals what is expected.
  console.error(`[catalog] UNHANDLED ${method} ${normalized}${query} — returning empty 200`)
  if (wantsAtom(req)) {
    sendAtom(res, req, [])
  } else {
    sendJson(res, { value: [] })
  }
  return true
}
