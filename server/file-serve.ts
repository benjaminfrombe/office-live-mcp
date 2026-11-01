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

/** Absolute image paths mentioned in tool output text, deduplicated in order. */
export function imagePathsInText(text: string): string[] {
  const matches = text.match(/(?:~?\/)[\w.@/-]+\.(?:png|jpe?g|gif|webp|svg)/gi) || []
  const seen = new Set<string>()
  const paths: string[] = []
  for (const match of matches) {
    // Drop relative paths: "./out.jpg" matches with its "./" as the slash.
    const start = text.indexOf(match)
    if (match.startsWith('/') && text[start - 1] === '.') continue
    if (!seen.has(match)) {
      seen.add(match)
      paths.push(match)
    }
  }
  return paths
}
