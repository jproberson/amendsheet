import { type Zippable, unzipSync, zipSync } from 'fflate'
import { XlsxError } from './errors.js'

/** Every part of the file, including the ones nothing here interprets. */
export interface Container {
  readonly parts: ReadonlyMap<string, Uint8Array>
}

export function readContainer(bytes: Uint8Array): Container {
  let entries: Record<string, Uint8Array>
  try {
    entries = unzipSync(bytes)
  } catch (cause) {
    throw new XlsxError('File is not a zip archive, so it cannot be an .xlsx file', { cause })
  }

  const parts = new Map<string, Uint8Array>()
  for (const [path, content] of Object.entries(entries)) {
    if (path.endsWith('/')) continue
    parts.set(path, content)
  }

  return { parts }
}

/** The ZIP epoch. A fixed stamp keeps the same parts producing the same bytes. */
const FIXED_TIMESTAMP = new Date(1980, 0, 1)

export function writeContainer(container: Container): Uint8Array {
  const entries: Zippable = {}
  for (const [path, content] of container.parts) {
    entries[path] = [content, { mtime: FIXED_TIMESTAMP }]
  }
  return zipSync(entries, { level: 6 })
}
