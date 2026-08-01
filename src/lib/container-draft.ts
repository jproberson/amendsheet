import { type Container, decodeXmlPart } from './container.js'
import { XlsxError } from './errors.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { relationshipsPathFor } from './workbook-parts.js'

export const EMPTY_RELATIONSHIPS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

/**
 * Adds a relationship to a rels part, opening a fresh one when there is none.
 * Returns the id it assigned as well as the text, since a part wired this way is
 * often referenced back by that id from the sheet — a legacy drawing is.
 */
export function withRelationship(
  relsXml: string | undefined,
  type: string,
  target: string,
): { xml: string; id: string } {
  const existing = relsXml ?? EMPTY_RELATIONSHIPS
  let maxId = 0
  for (const match of existing.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, Number(match[1]))
  const id = `rId${maxId + 1}`
  const relationship = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`
  const close = existing.indexOf('</Relationships>')
  if (close === -1) {
    throw new XlsxError('invalid-content', 'A relationships part is malformed', {})
  }
  return { xml: existing.slice(0, close) + relationship + existing.slice(close), id }
}

/**
 * A batch of part edits staged over a read container while a workbook is written.
 * The underlying map is the one `toBytes` hands the container: `null` deletes a
 * part, bytes replace it, everything not named passes through untouched.
 *
 * A part-level feature (comments, images, tables) reads the current text of a
 * part — its own pending write if it has one, else the original — stages new
 * bytes, and claims an unused part number, without knowing how the two stores
 * are layered. The primitives here are the ones every such feature repeated by
 * hand before.
 */
export interface ContainerDraft {
  /** The part's current text: a pending change if one is staged, else the
   * original. A part staged for deletion reads as its original, matching the
   * write pass, which only stages a delete for a part it then rewrites. */
  text(path: string): string | undefined
  /** The original part text, ignoring any pending change. */
  original(path: string): string | undefined
  setBytes(path: string, bytes: Uint8Array): void
  /** Whether the container holds this part or a change already stages it — the
   * test for skipping a number already taken when naming a fresh part. */
  has(path: string): boolean
  /** The smallest `n` at or above 1 whose `pathFor(n)` is not already taken.
   * Because a caller stages the part it claims before claiming the next, two
   * claims in a row return consecutive numbers. */
  freeNumber(pathFor: (n: number) => string): number
  /** The resolved target of `ownerPath`'s relationship of `type`, read from the
   * owner's current rels, or undefined when it declares none. */
  relationshipTarget(ownerPath: string, type: string): string | undefined
}

export function createContainerDraft(
  container: Container,
  changes: Map<string, Uint8Array | null>,
): ContainerDraft {
  const decoder = new TextDecoder()

  const original = (path: string): string | undefined => {
    const bytes = container.parts.get(path)
    return bytes === undefined ? undefined : decodeXmlPart(bytes, path)
  }

  const text = (path: string): string | undefined => {
    const changed = changes.get(path)
    if (changed !== undefined && changed !== null) return decoder.decode(changed)
    return original(path)
  }

  const has = (path: string): boolean => container.parts.has(path) || changes.has(path)

  return {
    text,
    original,
    has,
    setBytes(path, bytes) {
      changes.set(path, bytes)
    },
    freeNumber(pathFor) {
      let n = 0
      do {
        n += 1
      } while (has(pathFor(n)))
      return n
    },
    relationshipTarget(ownerPath, type) {
      const relsXml = text(relationshipsPathFor(ownerPath))
      if (relsXml === undefined) return undefined
      for (const relationship of readRelationships(relsXml, ownerPath).values()) {
        if (relationship.type === type && !relationship.external) {
          return resolveTarget(ownerPath, relationship.target)
        }
      }
      return undefined
    },
  }
}
