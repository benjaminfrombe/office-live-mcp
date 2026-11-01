import { describe, expect, it } from 'vitest'
import { imagePathsInText, resolveImageFileRequest } from './file-serve.ts'

describe('resolveImageFileRequest', () => {
  it('accepts an image path with its mime', () => {
    expect(resolveImageFileRequest('/tmp/render/slide-1.png')).toEqual({
      filePath: '/tmp/render/slide-1.png',
      mime: 'image/png',
    })
    const jpg = resolveImageFileRequest('/tmp/photo.JPG')
    expect(typeof jpg === 'object' && jpg.mime).toBe('image/jpeg')
  })

  it('rejects non-images and missing input', () => {
    expect(resolveImageFileRequest(null)).toBe('missing path parameter')
    expect(resolveImageFileRequest('/tmp/notes.md')).toBe('not an image file')
    expect(resolveImageFileRequest('/tmp/noextension')).toBe('not an image file')
    expect(resolveImageFileRequest('\0/etc/passwd.png')).toBe('invalid path')
  })
})

describe('imagePathsInText', () => {
  it('extracts absolute image paths from tool output', () => {
    const text = 'rendered to /tmp/checks/slide-3.png and /tmp/checks/overview.webp; see log.txt for details'
    expect(imagePathsInText(text)).toEqual(['/tmp/checks/slide-3.png', '/tmp/checks/overview.webp'])
  })

  it('supports ~/ paths and deduplicates', () => {
    const text = 'wrote ~/slide.png then wrote ~/slide.png again'
    expect(imagePathsInText(text)).toEqual(['~/slide.png'])
  })

  it('ignores relative paths and non-image extensions', () => {
    expect(imagePathsInText('saved slide.png and notes.md and ./local/out.jpg')).toEqual([])
  })
})
