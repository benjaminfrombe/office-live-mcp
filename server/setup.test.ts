import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  catalogUncCandidates,
  ensureLocalhostCert,
  harnessInstructions,
  isUncPath,
  MAC_APPS,
  manifestNameFor,
  mcpConfigSnippet,
  mergeDaemonTls,
  registryCommands,
} from './setup.ts'

afterEach(() => {
  delete process.env.OLM_CATALOG_UNC
})

describe('manifestNameFor', () => {
  it('picks the -https variant for TLS', () => {
    expect(manifestNameFor('manifest.xml', true)).toBe('manifest-https.xml')
    expect(manifestNameFor('manifest.word.xml', true)).toBe('manifest.word-https.xml')
  })

  it('keeps the plain manifest for HTTP', () => {
    expect(manifestNameFor('manifest.xml', false)).toBe('manifest.xml')
  })
})

describe('catalogUncCandidates', () => {
  it('puts an explicit override first', () => {
    process.env.OLM_CATALOG_UNC = '\\\\fileserver\\addin'
    const candidates = catalogUncCandidates('C:\\Users\\Admin\\AppData\\Local\\office-live-mcp', 'C:\\pkg')
    expect(candidates[0]).toBe('\\\\fileserver\\addin')
  })

  it('detects a Parallels host-share checkout', () => {
    const candidates = catalogUncCandidates('C:\\install', '\\\\Mac\\Home\\office-live-mcp')
    expect(candidates).toContain('\\\\Mac\\Home\\office-live-mcp\\addin')
  })

  it('falls back to a local share and the admin share', () => {
    const installDir = 'C:\\Users\\Admin\\AppData\\Local\\office-live-mcp'
    const candidates = catalogUncCandidates(installDir, 'C:\\pkg')
    expect(candidates.some((c) => c.startsWith('\\\\') && c.includes('OfficeLiveMCP'))).toBe(true)
    expect(candidates).toContain(`\\\\localhost\\C$\\Users\\Admin\\AppData\\Local\\office-live-mcp\\addin`)
  })
})

describe('mergeDaemonTls', () => {
  it('adds BRIDGE_TLS=1 to an empty env file', () => {
    expect(mergeDaemonTls('')).toBe('BRIDGE_TLS=1\n')
  })

  it('preserves other keys and comments, removes stale BRIDGE_TLS values', () => {
    const merged = mergeDaemonTls('# comment\nBRIDGE_PORT=9090\nBRIDGE_TLS=0\n')
    expect(merged).toContain('BRIDGE_PORT=9090')
    expect(merged).toContain('# comment')
    expect(merged).toContain('BRIDGE_TLS=1')
    expect(merged).not.toContain('BRIDGE_TLS=0')
  })

  it('is idempotent', () => {
    const once = mergeDaemonTls('OTHER=1\n')
    expect(mergeDaemonTls(once)).toBe(once)
  })
})

describe('ensureLocalhostCert', () => {
  it('generates a cert and key when absent', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'olm-cert-'))
    process.env.BRIDGE_CERT_PATH = join(dir, 'localhost.pem')
    process.env.BRIDGE_KEY_PATH = join(dir, 'localhost-key.pem')
    try {
      expect(await ensureLocalhostCert()).toBe(true)
      const cert = readFileSync(join(dir, 'localhost.pem'), 'utf8')
      expect(cert).toContain('BEGIN CERTIFICATE')
      expect(existsSync(join(dir, 'localhost-key.pem'))).toBe(true)
      // Idempotent: existing certs are kept.
      expect(await ensureLocalhostCert()).toBe(true)
    } finally {
      delete process.env.BRIDGE_CERT_PATH
      delete process.env.BRIDGE_KEY_PATH
      rmSync(dir, { recursive: true, force: true })
    }
  })
})

describe('harnessInstructions', () => {
  it('covers all three harnesses with the entry path filled in', () => {
    const lines = harnessInstructions('/opt/office-live-mcp/dist/index.cjs', 'darwin')
    const text = lines.join('\n')
    expect(text).toContain('OpenCode')
    expect(text).toContain('~/.config/opencode/opencode.jsonc')
    expect(text).toContain('claude mcp add office-live-mcp --scope user -- node "/opt/office-live-mcp/dist/index.cjs"')
    expect(text).toContain('Claude Desktop')
    expect(text).toContain('~/Library/Application Support/Claude/claude_desktop_config.json')
  })

  it('uses Windows config locations on win32', () => {
    const text = harnessInstructions('C:\\office-live-mcp\\dist\\index.cjs', 'win32').join('\n')
    expect(text).toContain('%USERPROFILE%\\.config\\opencode\\opencode.jsonc')
    expect(text).toContain('%APPDATA%\\Claude\\claude_desktop_config.json')
  })
})

describe('isUncPath', () => {
  it('accepts real UNC paths', () => {
    expect(isUncPath('\\\\Mac\\Home\\office-live-mcp\\addin')).toBe(true)
    expect(isUncPath('\\\\localhost\\C$\\Users\\a\\addin')).toBe(true)
  })

  it('rejects local paths and malformed values', () => {
    expect(isUncPath('C:\\o365-bridge\\addin')).toBe(false)
    expect(isUncPath('\\Mac\\Home\\addin')).toBe(false)
    expect(isUncPath('office-live-mcp')).toBe(false)
  })
})

describe('registryCommands', () => {
  it('registers Id, Url and Flags idempotently', () => {
    const commands = registryCommands('\\\\Mac\\Home\\office-live-mcp\\addin')
    expect(commands).toHaveLength(3)
    const flat = commands.map(([, args]) => args.join(' '))
    expect(flat[0]).toContain('TrustedCatalogs')
    expect(flat[1]).toContain('Url')
    expect(flat[1]).toContain('\\\\Mac\\Home\\office-live-mcp\\addin')
    expect(flat[2]).toContain('Flags')
    expect(flat.every((c) => c.includes('/f'))).toBe(true)
  })
})

describe('mcpConfigSnippet', () => {
  it('renders a valid client config with the entry command', () => {
    const snippet = mcpConfigSnippet('/usr/local/lib/node_modules/office-live-mcp/dist/index.js')
    const parsed = JSON.parse(snippet)
    expect(parsed.mcpServers['office-live-mcp'].command).toBe(
      '/usr/local/lib/node_modules/office-live-mcp/dist/index.js',
    )
  })
})

describe('MAC_APPS', () => {
  it('maps every app to a container and a manifest', () => {
    for (const app of Object.values(MAC_APPS)) {
      expect(app.container).toMatch(/^com\.microsoft\./)
      expect(app.manifest).toMatch(/^manifest(\.word|\.excel)?\.xml$/)
    }
  })
})
