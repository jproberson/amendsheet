import { type ShiftSpec, shiftFormula } from './shift.js'
import { readXml, withAttribute } from './xml.js'

interface Splice {
  readonly start: number
  readonly end: number
  readonly text: string
}

// Rewrites one attribute's value in place, keeping the tag's other bytes exact.
const mapAttribute = (tag: string, name: string, map: (value: string) => string): string =>
  tag.replace(
    new RegExp(`(\\s${name}\\s*=\\s*)("[^"]*"|'[^']*')`),
    (_, head: string, quoted: string) =>
      `${head}${quoted[0]}${map(quoted.slice(1, -1))}${quoted[0]}`,
  )

const encodeText = (text: string): string =>
  text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')

const applySplices = (xml: string, splices: readonly Splice[]): string => {
  let out = ''
  let position = 0
  for (const splice of [...splices].sort((a, b) => a.start - b.start)) {
    out += xml.slice(position, splice.start) + splice.text
    position = splice.end
  }
  return out + xml.slice(position)
}

// Attributes whose value is a reference or a space-separated list of them, each
// local to this sheet, so an inserted or deleted line shifts them like a formula.
const REFERENCE_ATTRIBUTE = new Map([
  ['dimension', 'ref'],
  ['mergeCell', 'ref'],
  ['autoFilter', 'ref'],
  ['conditionalFormatting', 'sqref'],
  ['dataValidation', 'sqref'],
  ['hyperlink', 'ref'],
])

// An inserted line is a monotonic shift, so its new position never lands before
// where it started; a column's `<col>` bounds and a row's number move the same way.
const shiftedIndex = (value: number, spec: ShiftSpec): number =>
  value < spec.at ? value : value + spec.delta

/**
 * Renumbers the edited sheet's cells and the rows or columns that carry them, and
 * shifts the references it holds — formula text, a shared formula's `ref`, merges,
 * the dimension, filters and the conditional-format, validation and hyperlink
 * ranges. Inserting is a monotonic shift, so no reference is destroyed.
 */
export function shiftSheet(xml: string, spec: ShiftSpec): string {
  const splices: Splice[] = []
  let rowNumber = 0
  let inFormula = false
  // While a deleted row is being skipped, the offset its removal splice starts at.
  let removingFrom = -1

  const spliceAttribute = (
    event: { start: number; end: number },
    name: string,
    map: (value: string) => string,
  ) => {
    const tag = xml.slice(event.start, event.end)
    const rewritten = mapAttribute(tag, name, map)
    if (rewritten !== tag) splices.push({ start: event.start, end: event.end, text: rewritten })
  }

  for (const event of readXml(xml)) {
    // A deleted row takes its cells, formulas and text with it, so everything up
    // to its close is dropped in one splice rather than shifted.
    if (removingFrom !== -1) {
      if (event.kind === 'close' && event.localName === 'row') {
        splices.push({ start: removingFrom, end: event.end, text: '' })
        removingFrom = -1
      }
      continue
    }
    if (event.kind === 'close') {
      if (event.localName === 'f') inFormula = false
      continue
    }
    if (event.kind === 'text') {
      if (!inFormula) continue
      const shifted = shiftFormula(event.text, spec)
      if (shifted !== event.text)
        splices.push({ start: event.start, end: event.end, text: encodeText(shifted) })
      continue
    }

    if (event.localName === 'row' && spec.axis === 'row') {
      const declared = event.attributes.get('r')
      rowNumber = declared === undefined ? rowNumber + 1 : Number(declared)
      // A row inside a deletion is dropped: a self-closing one at once, otherwise
      // from here to its close.
      if (spec.delta < 0 && rowNumber >= spec.at && rowNumber < spec.at - spec.delta) {
        if (event.selfClosing) splices.push({ start: event.start, end: event.end, text: '' })
        else removingFrom = event.start
        continue
      }
      // A row past the point moves; withAttribute makes its number explicit even
      // when the row left it implicit, so an inserted gap cannot renumber it wrong.
      const moved = shiftedIndex(rowNumber, spec)
      if (moved !== rowNumber) {
        const tag = xml.slice(event.start, event.end)
        splices.push({ start: event.start, end: event.end, text: withAttribute(tag, 'r', moved) })
      }
      continue
    }
    // A cols entry bounds a span of columns; both ends move, and a span the point
    // falls inside grows to cover the inserted column, keeping its left neighbour.
    if (event.localName === 'col' && spec.axis === 'column') {
      const tag = xml.slice(event.start, event.end)
      const bound = (value: string) => String(shiftedIndex(Number(value), spec))
      const rewritten = mapAttribute(mapAttribute(tag, 'min', bound), 'max', bound)
      if (rewritten !== tag) splices.push({ start: event.start, end: event.end, text: rewritten })
      continue
    }
    if (event.localName === 'c') {
      spliceAttribute(event, 'r', (value) => shiftFormula(value, spec))
      continue
    }
    if (event.localName === 'f') {
      if (!event.selfClosing) inFormula = true
      spliceAttribute(event, 'ref', (value) => shiftFormula(value, spec))
      continue
    }
    const referenceAttribute = REFERENCE_ATTRIBUTE.get(event.localName)
    if (referenceAttribute !== undefined)
      spliceAttribute(event, referenceAttribute, (value) => shiftFormula(value, spec))
  }

  return applySplices(xml, splices)
}

/**
 * Rewrites only the formula text of another sheet, where a reference qualified
 * with the edited sheet points into the rows or columns that moved. The sheet's
 * own rows, cells and ranges are local to it and stay exactly as they were.
 */
export function shiftForeignFormulas(xml: string, spec: ShiftSpec): string {
  const splices: Splice[] = []
  let inFormula = false
  for (const event of readXml(xml)) {
    if (event.kind === 'close') {
      if (event.localName === 'f') inFormula = false
      continue
    }
    if (event.kind === 'text') {
      if (!inFormula) continue
      const shifted = shiftFormula(event.text, spec)
      if (shifted !== event.text)
        splices.push({ start: event.start, end: event.end, text: encodeText(shifted) })
      continue
    }
    if (event.localName === 'f' && !event.selfClosing) inFormula = true
  }
  return applySplices(xml, splices)
}
