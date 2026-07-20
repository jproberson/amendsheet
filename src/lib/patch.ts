import { dateToSerial } from './date.js'
import { XlsxError } from './errors.js'
import { readXml } from './xml.js'

/** What a caller may put in a cell. */
export type CellInput = number | string | boolean | Date | null

const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/**
 * Builds the replacement element for one cell. The style attribute is carried
 * over from the cell being replaced so formatting is not lost, which also means
 * writing a Date into a cell that has no date format will show as a number.
 */
function cellElement(
  reference: string,
  value: CellInput,
  style: string | undefined,
  date1904: boolean,
): string {
  const attributes = style === undefined ? '' : ` s="${style}"`

  if (value === null) return `<c r="${reference}"${attributes}/>`

  if (typeof value === 'string') {
    return `<c r="${reference}"${attributes} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`
  }
  if (typeof value === 'boolean') {
    return `<c r="${reference}"${attributes} t="b"><v>${value ? 1 : 0}</v></c>`
  }
  if (value instanceof Date) {
    return `<c r="${reference}"${attributes}><v>${dateToSerial(value, date1904)}</v></c>`
  }
  if (!Number.isFinite(value)) {
    throw new XlsxError(`Cell ${reference} cannot hold ${value}`)
  }
  return `<c r="${reference}"${attributes}><v>${value}</v></c>`
}

/**
 * Rewrites only the cells named in `edits`, copying every other byte of the
 * sheet through unchanged. Row attributes, merged cells, conditional formatting
 * and anything else this library does not model survive because they are never
 * re-serialised.
 */
export function patchSheet(
  xml: string,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
): string {
  if (edits.size === 0) return xml

  const pieces: string[] = []
  const applied = new Set<string>()
  let cursor = 0

  let reference: string | undefined
  let style: string | undefined
  let elementStart = 0

  const replace = (from: number, to: number, name: string) => {
    pieces.push(xml.slice(cursor, from))
    pieces.push(cellElement(name, edits.get(name) ?? null, style, date1904))
    cursor = to
    applied.add(name)
  }

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.name === 'c') {
      reference = event.attributes.get('r')
      style = event.attributes.get('s')
      elementStart = event.start

      if (event.selfClosing && reference !== undefined && edits.has(reference)) {
        replace(event.start, event.end, reference)
      }
      continue
    }

    if (event.kind === 'close' && event.name === 'c') {
      if (reference !== undefined && edits.has(reference) && !applied.has(reference)) {
        replace(elementStart, event.end, reference)
      }
      reference = undefined
    }
  }

  pieces.push(xml.slice(cursor))

  for (const name of edits.keys()) {
    if (!applied.has(name)) {
      throw new XlsxError(`Cell ${name} is not in this sheet; adding cells is not supported yet`)
    }
  }

  return pieces.join('')
}
