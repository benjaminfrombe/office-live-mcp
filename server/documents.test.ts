import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { validateDocumentPath } from './tools/documents.ts'

describe('open_document validation', () => {
  it('accepts an existing .pptx file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'olm-open-'))
    const file = join(dir, 'deck.pptx')
    writeFileSync(file, 'stub')
    expect(validateDocumentPath(file)).toEqual({ ext: '.pptx' })
  })

  it('rejects a missing file', () => {
    expect(() => validateDocumentPath('/nonexistent/deck.pptx')).toThrow(/File not found/)
  })

  it('rejects unsupported extensions', () => {
    const dir = mkdtempSync(join(tmpdir(), 'olm-open-'))
    const file = join(dir, 'script.sh')
    writeFileSync(file, 'stub')
    expect(() => validateDocumentPath(file)).toThrow(/Unsupported file type/)
  })

  it('rejects an empty path', () => {
    expect(() => validateDocumentPath('')).toThrow(/path is required/)
  })
})
