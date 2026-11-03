// GET /file?path=... — serve local image files (tool outputs like rendered
// slide checks) to the add-in. Validation is pure so it is unit-testable:
// only absolute paths with an image extension may be served, with a fixed
// mime map. The bridge binds 127.0.0.1, so this never leaves the machine.

var IMAGE_MIME_BY_EXTENSION: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
}

export interface ImageFileRequest {
  filePath: string
  mime: string
}

/** Resolve a /file?path= request to a servable image file, or an error string. */
export function resolveImageFileRequest(rawPath: string | null): ImageFileRequest | string {
  if (!rawPath) return 'missing path parameter'
  if (rawPath.startsWith('\0')) return 'invalid path'
  const dot = rawPath.lastIndexOf('.')
  if (dot === -1) return 'not an image file'
  const mime = IMAGE_MIME_BY_EXTENSION[rawPath.slice(dot).toLowerCase()]
  if (!mime) return 'not an image file'
  return { filePath: rawPath, mime }
}

/** Absolute image paths mentioned in tool output text, deduplicated in order.
 * Tools often wrap exact paths in <file>...</file> tags (spaces included);
 * bare mentions fall back to a space-free path regex. */
export function imagePathsInText(text: string): string[] {
  const paths: string[] = []
  const seen = new Set<string>()
  const push = (candidate: string) => {
    const dot = candidate.lastIndexOf('.')
    if (dot === -1) return
    if (!IMAGE_MIME_BY_EXTENSION[candidate.slice(dot).toLowerCase()]) return
    if (!candidate.startsWith('/') && !candidate.startsWith('~/')) return
    // Drop relative paths: "./out.jpg" matches with its "./" as the slash.
    const start = text.indexOf(candidate)
    if (candidate.startsWith('/') && text[start - 1] === '.') return
    if (!seen.has(candidate)) {
      // A bare match can be the tail of an already-extracted <file> path
      // (the regex restarts after a space) — never push it as a second image.
      if ([...seen].some((existing) => existing.endsWith(candidate))) return
      seen.add(candidate)
      paths.push(candidate)
    }
  }
  // Exact paths wrapped in <file>...</file> — may contain spaces.
  const wrapped = text.match(/<file>([^<]+)<\/file>/g) || []
  for (const wrapper of wrapped) {
    push(wrapper.slice(6, -7).trim())
  }
  // Bare mentions: space-free absolute paths.
  const matches = text.match(/(?:~?\/)[\w.@/-]+/g) || []
  for (const match of matches) {
    push(match)
  }
  return paths
}
