#!/usr/bin/env node

/**
 * Sideloads the add-in manifest into an Office app's WEF directory,
 * substituting the correct port from BRIDGE_PORT env var.
 *
 * Usage:
 *   node --experimental-strip-types scripts/sideload.ts                    # PowerPoint (HTTP), default port 8790
 *   node --experimental-strip-types scripts/sideload.ts --app word       # Word
 *   node --experimental-strip-types scripts/sideload.ts --app excel      # Excel
 *   node --experimental-strip-types scripts/sideload.ts --tls            # HTTPS manifest, default port 8791
 *   BRIDGE_PORT=9090 node --experimental-strip-types scripts/sideload.ts # Custom port
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { substituteManifestPort } from '../server/manifest.ts'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const PROJECT_ROOT = resolve(SCRIPT_DIR, '..')

const APPS = {
  powerpoint: { manifest: 'manifest.xml', container: 'com.microsoft.Powerpoint' },
  word: { manifest: 'manifest.word.xml', container: 'com.microsoft.Word' },
  excel: { manifest: 'manifest.excel.xml', container: 'com.microsoft.Excel' },
} as const
type AppKey = keyof typeof APPS

function parseApp(): AppKey {
  const idx = process.argv.indexOf('--app')
  const value = idx !== -1 ? process.argv[idx + 1] : 'powerpoint'
  if (value === 'powerpoint' || value === 'word' || value === 'excel') return value
  throw new Error(`Unknown app "${value}". Valid values: ${Object.keys(APPS).join(', ')}`)
}

const app = parseApp()
const tls = process.argv.includes('--tls')
const defaultPort = tls ? 8791 : 8790
const port = Number(process.env.BRIDGE_PORT) || defaultPort

const manifestName = tls ? APPS[app].manifest.replace('.xml', '-https.xml') : APPS[app].manifest
const src = resolve(PROJECT_ROOT, 'addin', manifestName)
const wefDir = resolve(homedir(), `Library/Containers/${APPS[app].container}/Data/Documents/wef`)
const dest = resolve(wefDir, 'manifest.xml')

// Read template and substitute port (shared with the bridge server via substituteManifestPort)
const template = readFileSync(src, 'utf8')
const content = substituteManifestPort(template, defaultPort, port)

// Write to WEF directory
mkdirSync(wefDir, { recursive: true })
writeFileSync(dest, content)

// Write version marker
const pkg = JSON.parse(readFileSync(resolve(PROJECT_ROOT, 'package.json'), 'utf8'))
writeFileSync(resolve(PROJECT_ROOT, '.sideloaded'), `${pkg.version}:${port}`)

console.log(`[sideload] ${app} manifest installed (${tls ? 'HTTPS' : 'HTTP'}, port ${port})`)
