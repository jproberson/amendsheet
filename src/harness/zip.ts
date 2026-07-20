import { unzipSync, zipSync } from 'fflate'
import type { Parts } from './types.js'

/**
 * An .xlsx file is a ZIP of XML parts. Reading the raw part list is how we
 * detect data loss that a library's own object model would hide from us:
 * if a chart part goes in and does not come out, the document was damaged
 * no matter what the cell values say.
 */
export function readParts(bytes: Uint8Array): Parts {
  return new Map(Object.entries(unzipSync(bytes)))
}

export function writeParts(parts: Record<string, Uint8Array>): Uint8Array {
  return zipSync(parts, { level: 6 })
}

export function decodeText(bytes: Uint8Array): string {
  return new TextDecoder('utf-8').decode(bytes)
}

export function encodeText(text: string): Uint8Array {
  return new TextEncoder().encode(text)
}

/**
 * Copy a view into a standalone ArrayBuffer sized to exactly its contents.
 * Copying rather than slicing the backing store keeps the return type honest:
 * a Uint8Array may be backed by a SharedArrayBuffer.
 */
export function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new ArrayBuffer(bytes.byteLength)
  new Uint8Array(copy).set(bytes)
  return copy
}

export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false
  }
  return true
}
