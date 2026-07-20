import { unzipSync, zipSync } from 'fflate'
import { XlsxError } from './errors.js'

/**
 * An .xlsx file is a ZIP of XML parts. Holding all of them — including the ones
 * nothing here understands — is what lets a document survive a round trip with
 * its charts, pivot tables and drawings intact.
 */
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
    // Directory entries carry no content and are recreated on write.
    if (path.endsWith('/')) continue
    parts.set(path, content)
  }

  return { parts }
}

export function writeContainer(container: Container): Uint8Array {
  const entries: Record<string, Uint8Array> = {}
  for (const [path, content] of container.parts) {
    entries[path] = content
  }
  return zipSync(entries, { level: 6 })
}
