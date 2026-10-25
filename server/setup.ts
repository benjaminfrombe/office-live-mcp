// `office-live-mcp setup` — one command, per platform:
//
//   macOS:  sideload the add-in manifests into each installed Office app's
//           WEF directory (HTTP or HTTPS depending on daemon.env), then print
//           ready-to-paste MCP client config.
//   Windows: copy dist/ + addin/ to a fixed install dir, register that folder
//           as a trusted add-in catalog (HKCU registry, no admin needed), and
//           print the MCP config plus the one step no installer can do:
//           restart Office → Home → Add-ins → SHARED FOLDER → Add.
//
// Everything the setup does is idempotent: running it twice is safe.

import { execFileSync, spawnSync } from 'node:child_process'
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir, hostname } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import selfsigned from 'selfsigned'
import { substituteManifestPort } from './manifest.ts'

const SCRIPT_DIR = typeof __dirname !== 'undefined' ? __dirname : dirname(fileURLToPath(import.meta.url))
const PROJECT_ROOT = resolve(SCRIPT_DIR, '..')
const CATALOG_GUID = '{7E3F2A11-93C4-4B5D-8E6F-1A2B3C4D5E6F}'
const REG_KEY = `HKCU\\Software\\Microsoft\\Office\\16.0\\Wef\\TrustedCatalogs\\${CATALOG_GUID}`
const CERT_PATH = () => process.env.BRIDGE_CERT_PATH || resolve(PROJECT_ROOT, 'certs', 'localhost.pem')
const CERT_KEY_PATH = () => process.env.BRIDGE_KEY_PATH || resolve(PROJECT_ROOT, 'certs', 'localhost-key.pem')
const HTTPS_CATALOG_URL = 'https://localhost:8443'

// ---------------------------------------------------------------------------
// Pure helpers (unit-tested)
// ---------------------------------------------------------------------------

export interface MacAppDef {
  manifest: string
  container: string
  label: string
}

export const MAC_APPS: Record<'powerpoint' | 'word' | 'excel', MacAppDef> = {
  powerpoint: { manifest: 'manifest.xml', container: 'com.microsoft.Powerpoint', label: 'PowerPoint' },
  word: { manifest: 'manifest.word.xml', container: 'com.microsoft.Word', label: 'Word' },
  excel: { manifest: 'manifest.excel.xml', container: 'com.microsoft.Excel', label: 'Excel' },
}

/** The manifest file variant for the transport the daemon will use. */
export function manifestNameFor(manifest: string, tls: boolean): string {
  return tls ? manifest.replace('.xml', '-https.xml') : manifest
}

/**
 * Decide the UNC catalog path for the add-in folder on Windows. Candidates in
 * priority order; the first one that exists (checked by the caller) wins.
 * Exported for tests.
 */
export function catalogUncCandidates(installDir: string, packageRoot: string): string[] {
  const candidates: string[] = []
  // 1. Explicit override.
  if (process.env.OLM_CATALOG_UNC) candidates.push(process.env.OLM_CATALOG_UNC)
  // 2. Running straight from the Parallels host share — the addin folder is
  //    already reachable as a UNC path, no share needed.
  if (packageRoot.startsWith('\\\\Mac\\Home\\')) candidates.push(`${packageRoot}\\addin`)
  // 3. A dedicated local share (created by setup when elevation is available).
  candidates.push(`\\\\${hostname()}\\OfficeLiveMCP`)
  // 4. The default admin share — exists on Windows Pro for admin users.
  const drive = installDir.charAt(0)
  if (installDir.startsWith('C:\\Users\\')) {
    candidates.push(`\\\\localhost\\${drive}$\\${installDir.slice(3)}\\addin`)
  }
  return candidates
}

/** True when the path is a usable UNC path (Office only trusts network paths). */
export function isUncPath(path: string): boolean {
  return /^\\\\[^\\]+\\.+/.test(path)
}

/** reg.exe commands that register the catalog (idempotent: /f overwrites). */
export function registryCommands(uncPath: string): [string, string[]][] {
  return [
    ['reg', ['add', REG_KEY, '/v', 'Id', '/t', 'REG_SZ', '/d', CATALOG_GUID, '/f']],
    ['reg', ['add', REG_KEY, '/v', 'Url', '/t', 'REG_SZ', '/d', uncPath, '/f']],
    ['reg', ['add', REG_KEY, '/v', 'Flags', '/t', 'REG_DWORD', '/d', '1', '/f']],
  ]
}

/** Ready-to-paste MCP client config for the given server entry. */
export function mcpConfigSnippet(entryCommand: string): string {
  const command = JSON.stringify(entryCommand)
  return ['{', `  "mcpServers": {`, `    "office-live-mcp": { "command": ${command} }`, '  }', '}'].join('\n')
}

/**
 * Per-harness add instructions, with the real entry path filled in and
 * platform-aware config locations. Returned as lines for testability.
 * nodePath must be an absolute node executable (Windows users often have no
 * node on PATH, and a .cjs entry is not directly executable there).
 */
export function harnessInstructions(entryCommand: string, platform: NodeJS.Platform, nodePath: string): string[] {
  const isMac = platform === 'darwin'
  const lines: string[] = []

  lines.push('Add the MCP server to your AI harness (pick the one you use):')
  lines.push('')
  lines.push('  OpenCode')
  lines.push(
    `    Config: ${isMac ? '~/.config/opencode/opencode.jsonc' : '%USERPROFILE%\\.config\\opencode\\opencode.jsonc'}`,
  )
  lines.push('    Add to the "mcp" section:')
  lines.push('      {')
  lines.push('        "office-live-mcp": {')
  lines.push('          "type": "local",')
  lines.push(`          "command": ["${nodePath}", "${entryCommand}"],`)
  lines.push('          "enabled": true')
  lines.push('        }')
  lines.push('      }')
  lines.push('')
  lines.push('  Claude Code')
  lines.push('    Run:')
  lines.push(`      claude mcp add office-live-mcp --scope user -- ${nodePath} "${entryCommand}"`)
  lines.push('')
  lines.push('  Claude Desktop')
  lines.push(
    `    Config: ${
      isMac
        ? '~/Library/Application Support/Claude/claude_desktop_config.json'
        : '%APPDATA%\\Claude\\claude_desktop_config.json'
    }`,
  )
  lines.push('    Add to "mcpServers":')
  lines.push('      {')
  lines.push('        "office-live-mcp": {')
  lines.push(`          "command": "${nodePath}",`)
  lines.push(`          "args": ["${entryCommand}"]`)
  lines.push('        }')
  lines.push('      }')
  return lines
}

// ---------------------------------------------------------------------------
// macOS setup
// ---------------------------------------------------------------------------

function daemonTlsEnabled(): boolean {
  try {
    const envFile = resolve(homedir(), '.office-live-mcp', 'daemon.env')
    return /^BRIDGE_TLS=1\b/m.test(readFileSync(envFile, 'utf8'))
  } catch {
    return false
  }
}

function setupMac(): void {
  const tls = daemonTlsEnabled()
  const defaultPort = tls ? 8443 : 8080
  const port = Number(process.env.BRIDGE_PORT) || defaultPort
  let installed = 0
  let found = 0

  for (const [key, app] of Object.entries(MAC_APPS)) {
    const containerDir = resolve(homedir(), `Library/Containers/${app.container}`)
    if (!existsSync(containerDir)) {
      console.log(`[setup] ${app.label}: not installed, skipped`)
      continue
    }
    found++
    const manifest = manifestNameFor(app.manifest, tls)
    const src = resolve(PROJECT_ROOT, 'addin', manifest)
    if (!existsSync(src)) {
      console.log(`[setup] ${app.label}: manifest ${manifest} missing in package, skipped`)
      continue
    }
    const wefDir = resolve(containerDir, 'Data/Documents/wef')
    const template = readFileSync(src, 'utf8')
    try {
      mkdirSync(wefDir, { recursive: true })
      writeFileSync(resolve(wefDir, 'manifest.xml'), substituteManifestPort(template, defaultPort, port))
    } catch (err: unknown) {
      const code = (err as { code?: string }).code
      if (code === 'EPERM' || code === 'EACCES') {
        console.log(
          `[setup] ${app.label}: blocked by macOS privacy protection (TCC). ` +
            'Grant your terminal Full Disk Access (System Settings → Privacy & Security) and rerun setup, ' +
            `or run: node --experimental-strip-types scripts/sideload.ts --app ${key}`,
        )
        continue
      }
      throw err
    }
    console.log(`[setup] ${app.label}: add-in manifest installed (${tls ? 'HTTPS' : 'HTTP'}, port ${port})`)
    installed++
  }

  if (found === 0) {
    console.log('[setup] No Office apps found — nothing to sideload.')
  } else if (installed === 0) {
    console.log('[setup] No manifests were installed — see the messages above.')
  } else {
    console.log('[setup] Restart your Office apps, open a document, and click the add-in button in the Home ribbon.')
  }
}

// ---------------------------------------------------------------------------
// Windows setup
// ---------------------------------------------------------------------------

/** Generate a localhost TLS cert (selfsigned, valid 10 years) if absent. */
export async function ensureLocalhostCert(): Promise<boolean> {
  if (existsSync(CERT_PATH()) && existsSync(CERT_KEY_PATH())) {
    return true
  }
  try {
    const pems = await selfsigned.generate([{ name: 'commonName', value: 'localhost' }], {
      notAfterDate: new Date(Date.now() + 3650 * 24 * 3600 * 1000),
      keySize: 2048,
      algorithm: 'sha256',
      extensions: [
        {
          name: 'subjectAltName',
          altNames: [
            { type: 2, value: 'localhost' }, // DNS
            { type: 2, value: hostname().toLowerCase() }, // machine name — Office may
            // refuse loopback as a SharePoint-style catalog site
            { type: 7, ip: '127.0.0.1' }, // IP
            { type: 7, ip: '::1' },
          ],
        },
      ],
    })
    mkdirSync(dirname(CERT_PATH()), { recursive: true })
    writeFileSync(CERT_PATH(), pems.cert)
    writeFileSync(CERT_KEY_PATH(), pems.private)
    console.log(`[setup] Generated localhost TLS certificate: ${CERT_PATH()}`)
    return true
  } catch (err: unknown) {
    console.error(`[setup] Certificate generation failed: ${err instanceof Error ? err.message : String(err)}`)
    return false
  }
}

/** Merge BRIDGE_TLS=1 into existing daemon.env content (pure). */
export function mergeDaemonTls(content: string): string {
  if (/^BRIDGE_TLS=1\b/m.test(content)) return content
  const lines = content.split('\n').filter((l) => l.trim() && !/^BRIDGE_TLS=/.test(l))
  lines.push('BRIDGE_TLS=1')
  return `${lines.join('\n')}\n`
}

/** Merge BRIDGE_TLS=1 into the daemon env file. */
export function writeDaemonTlsEnv(): void {
  const envFile = resolve(homedir(), '.office-live-mcp', 'daemon.env')
  let content = ''
  try {
    content = readFileSync(envFile, 'utf8')
  } catch {
    // no file yet
  }
  const merged = mergeDaemonTls(content)
  if (merged !== content) {
    mkdirSync(dirname(envFile), { recursive: true })
    writeFileSync(envFile, merged)
  }
  console.log(`[setup] Daemon env: ${envFile} (BRIDGE_TLS=1 — bridge serves https://localhost:8443)`)
}

/**
 * Primary Windows route, no share involved: the bridge itself acts as an
 * HTTPS add-in catalog (SharePoint catalog emulation). Office enumerates
 * https://localhost:8443 and lists the add-in under MY ORGANIZATION.
 * Returns false when a step fails so the caller can fall back to UNC.
 */
async function setupWindowsHttpsCatalog(): Promise<boolean> {
  if (!(await ensureLocalhostCert())) return false

  // Trust the cert in the Windows Root store. Try the current-user store
  // first (no elevation; Windows shows a one-time confirmation dialog). If
  // that fails (e.g. running as SYSTEM, or the user cancelled), try the
  // local-machine store, which works silently from elevated contexts.
  const userStore = spawnSync('certutil', ['-user', '-addstore', 'Root', CERT_PATH()], { stdio: 'pipe' })
  let trusted = userStore.status === 0
  if (!trusted) {
    const machineStore = spawnSync('certutil', ['-addstore', 'Root', CERT_PATH()], { stdio: 'pipe' })
    trusted = machineStore.status === 0
  }
  if (!trusted) {
    console.error(
      '[setup] Could not trust the certificate in the Windows Root store (certutil failed or was cancelled).',
    )
    return false
  }
  console.log('[setup] Certificate trusted in the Windows Root store')

  writeDaemonTlsEnv()

  for (const [cmd, args] of registryCommands(HTTPS_CATALOG_URL)) {
    try {
      execFileSync(cmd, args, { stdio: 'pipe' })
    } catch (err: unknown) {
      console.error(`[setup] Failed to write registry key: ${err instanceof Error ? err.message : String(err)}`)
      return false
    }
  }
  console.log(`[setup] Trusted add-in catalog registered: ${HTTPS_CATALOG_URL} (registry, current user only)`)

  console.log(
    [
      '[setup] One manual step left (Office requires it):',
      '  1. Close and reopen PowerPoint / Word / Excel',
      '  2. Home ribbon → Add-ins → Advanced... → MY ORGANIZATION tab',
      '  3. Select "Office Live MCP" → Add',
    ].join('\n'),
  )
  return true
}

/** Fallback Windows route: shared folder catalog (UNC path, as Office documents). */
function setupWindowsUncCatalog(): void {
  const installDir = process.env.OLM_INSTALL_DIR || resolve(homedir(), 'AppData', 'Local', 'office-live-mcp')

  // 1. Copy dist + addin into the install dir.
  for (const folder of ['dist', 'addin']) {
    const src = resolve(PROJECT_ROOT, folder)
    if (!existsSync(src)) {
      console.error(`[setup] Package incomplete: ${src} is missing.`)
      process.exit(1)
    }
    cpSync(src, resolve(installDir, folder), { recursive: true })
  }
  // Prune the manifest variants that do not match the daemon transport, so
  // the catalog lists exactly one add-in per Office app. The catalog
  // experiment files (manifest.vm-*.xml) are never listed.
  const tls = daemonTlsEnabled()
  for (const file of readdirSync(resolve(installDir, 'addin'))) {
    if (!file.endsWith('.xml')) continue
    const isHttpsVariant = file.includes('-https')
    const isVmExperiment = file.includes('.vm-')
    if (isVmExperiment || (tls && !isHttpsVariant) || (!tls && isHttpsVariant)) {
      rmSync(resolve(installDir, 'addin', file))
    }
  }
  console.log(`[setup] Installed files to ${installDir} (manifests: ${tls ? 'HTTPS' : 'HTTP'})`)

  // 2. Determine a reachable UNC path for the addin catalog.
  const candidates = catalogUncCandidates(installDir, PROJECT_ROOT)
  let unc: string | undefined
  for (const candidate of candidates) {
    if (!isUncPath(candidate)) {
      console.error(`[setup] Ignoring candidate "${candidate}" — not a valid UNC path (expected \\\\host\\share).`)
      continue
    }
    if (existsSync(candidate)) {
      unc = candidate
      break
    }
  }
  if (!unc) {
    // Try to create the dedicated share. Needs elevation; failures fall
    // through to the admin-share candidate, which needs no elevation but
    // only exists for admin users on Pro/Enterprise.
    const addinDir = resolve(installDir, 'addin')
    const result = spawnSync('net', ['share', `OfficeLiveMCP=${addinDir}`, '/GRANT:Everyone:READ'], {
      stdio: 'pipe',
    })
    if (result.status === 0 && existsSync(`\\\\${hostname()}\\OfficeLiveMCP`)) {
      unc = `\\\\${hostname()}\\OfficeLiveMCP`
    }
  }
  if (!unc) {
    const adminShare = candidates.find((c) => c.startsWith('\\\\localhost\\'))
    if (adminShare && existsSync(adminShare)) {
      unc = adminShare
    }
  }
  if (!unc) {
    console.error(
      [
        '[setup] Could not find a UNC path for the add-in catalog (Office only trusts network paths).',
        '  Tried:',
        ...candidates.map((c) => `    - ${c}`),
        '  Fixes, pick one:',
        '    - Parallels: run setup from the host share (\\\\Mac\\Home\\...) or set OLM_CATALOG_UNC',
        '    - Native Windows: run setup from an elevated terminal so it can create the share,',
        '      or set OLM_CATALOG_UNC to an existing share that exposes the addin folder.',
      ].join('\n'),
    )
    process.exit(1)
  }
  console.log(`[setup] Add-in catalog: ${unc}`)

  // 3. Register the catalog (HKCU — no elevation needed).
  for (const [cmd, args] of registryCommands(unc)) {
    try {
      execFileSync(cmd, args, { stdio: 'pipe' })
    } catch (err: unknown) {
      console.error(`[setup] Failed to write registry key: ${err instanceof Error ? err.message : String(err)}`)
      process.exit(1)
    }
  }
  console.log('[setup] Trusted add-in catalog registered (registry, current user only)')

  console.log(
    [
      '[setup] One manual step left (Office requires it):',
      '  1. Close and reopen PowerPoint / Word / Excel',
      '  2. Home ribbon → Add-ins → Advanced... → SHARED FOLDER tab',
      '  3. Select "Office Live MCP" → Add',
    ].join('\n'),
  )
}

async function setupWindows(): Promise<void> {
  // Default: shared-folder route via a UNC path. On admin machines the
  // built-in admin share (\\localhost\C$) satisfies Office's UNC requirement
  // with zero user setup. The HTTPS SharePoint-catalog route (OLM_HTTPS_CATALOG=1)
  // is legacy: Microsoft fully retired SharePoint add-in catalogs on
  // April 2, 2026, and Office builds after that ignore the enumeration
  // results — verified live against Office 16.0.20527.
  if (process.env.OLM_HTTPS_CATALOG === '1') {
    console.log(
      '[setup] Warning: the HTTPS SharePoint-catalog route is legacy — Office builds after April 2026 ignore it.',
    )
    if (await setupWindowsHttpsCatalog()) return
    console.log('[setup] Falling back to the shared-folder (UNC) route.')
    setupWindowsUncCatalog()
    return
  }
  setupWindowsUncCatalog()
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export async function runSetup(): Promise<void> {
  console.log(`office-live-mcp setup (${process.platform})`)
  if (process.platform === 'darwin') {
    setupMac()
  } else if (process.platform === 'win32') {
    await setupWindows()
  } else {
    console.error(`[setup] Unsupported platform: ${process.platform}`)
    process.exit(1)
  }

  const entry = process.argv[1] ?? 'office-live-mcp'
  console.log(['', ...harnessInstructions(entry, process.platform, process.execPath)].join('\n'))
}
