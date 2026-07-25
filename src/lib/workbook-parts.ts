import { columnToIndex } from './reference.js'
import { resolveTarget } from './relationships.js'
import { type ShiftSpec, shiftFormula } from './shift.js'
import { readXml, readXmlBytes } from './xml.js'

/**
 * CT_Workbook is a sequence, so a calcPr the file lacks cannot simply be
 * appended: these are the children the schema puts after it.
 */
const AFTER_CALC_PR = new Set([
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
])

/**
 * A written formula carries no computed result, so the workbook is marked for
 * recalculation. Without it a reader that trusts cached values shows nothing.
 */
export function withRecalculation(xml: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'calcPr') continue

    const tag = xml.slice(event.start, event.end)
    if (tag.includes('fullCalcOnLoad=')) {
      return (
        xml.slice(0, event.start) +
        // Either quote character is legal, and matching only double quotes
        // left the check passing while the rewrite did nothing.
        tag.replace(/fullCalcOnLoad=("|')[^"']*\1/, 'fullCalcOnLoad="1"') +
        xml.slice(event.end)
      )
    }
    const opened = tag.replace(/\/?>$/, (end) =>
      end === '/>' ? ' fullCalcOnLoad="1"/>' : ' fullCalcOnLoad="1">',
    )
    return xml.slice(0, event.start) + opened + xml.slice(event.end)
  }

  let depth = 0
  let insertAt = -1

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      // Direct children of the root open at depth 1, so a nested extLst in a
      // part we do not interpret cannot be mistaken for the workbook's own.
      if (depth === 1 && insertAt === -1 && AFTER_CALC_PR.has(event.localName)) {
        insertAt = event.start
      }
      if (!event.selfClosing) depth++
      continue
    }
    if (event.kind !== 'close') continue

    depth--
    if (event.localName !== 'workbook' || depth !== 0) continue

    const colon = event.name.indexOf(':')
    const prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
    const element = `<${prefix}calcPr fullCalcOnLoad="1"/>`
    const at = insertAt === -1 ? event.start : insertAt
    return xml.slice(0, at) + element + xml.slice(at)
  }

  return xml
}

/**
 * Removes the relationship pointing at one part, leaving every other byte
 * alone. A relationship whose target is gone is an invalid package.
 */
export function withoutRelationship(xml: string, ownerPath: string, part: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Relationship') continue
    if (event.attributes.get('TargetMode') === 'External') continue
    const target = event.attributes.get('Target')
    if (target === undefined || resolveTarget(ownerPath, target) !== part) continue
    return xml.slice(0, event.start) + xml.slice(event.end)
  }
  return xml
}

/** Removes one Override element, leaving every other byte of the part alone. */
export function withoutOverride(xml: string, part: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Override') continue
    if (event.attributes.get('PartName') !== `/${part}`) continue
    return xml.slice(0, event.start) + xml.slice(event.end)
  }
  return xml
}

/** A part's relationships live in a `_rels` folder beside it, named after it. */
export const relationshipsPathFor = (partPath: string): string =>
  partPath.replace(/([^/]+)$/, '_rels/$1.rels')

/** The highest row a sheet stores, so a shift that would push one off is refused. */
export function highestRow(bytes: Uint8Array): number {
  let highest = 0
  let current = 0
  for (const event of readXmlBytes(bytes)) {
    if (event.kind === 'open' && event.localName === 'row') {
      const declared = event.attributes.get('r')
      current = declared === undefined ? current + 1 : Number(declared)
      highest = Math.max(highest, current)
    }
  }
  return highest
}

/** The highest column a sheet stores, so a shift that would push one off is refused. */
export function highestColumn(bytes: Uint8Array): number {
  let highest = 0
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'c') {
      const reference = event.attributes.get('r')
      const letters = reference?.match(/[A-Za-z]+/)?.[0]
      if (letters !== undefined) highest = Math.max(highest, columnToIndex(letters))
    } else if (event.localName === 'col') {
      const max = event.attributes.get('max')
      if (max !== undefined) highest = Math.max(highest, Number(max))
    }
  }
  return highest
}

// Elements whose reference a deletion could collapse to nothing. A formula's own
// text may become #REF! and still be valid, but a structural range or a shared
// formula's home cell cannot, so a deletion that would do it is refused instead.
const COLLAPSIBLE = new Map([
  ['mergeCell', { attribute: 'ref', reason: 'a merged range' }],
  ['autoFilter', { attribute: 'ref', reason: 'the filter' }],
  ['hyperlink', { attribute: 'ref', reason: 'a hyperlink' }],
  ['conditionalFormatting', { attribute: 'sqref', reason: 'a conditional format' }],
  ['dataValidation', { attribute: 'sqref', reason: 'a data validation' }],
  ['dimension', { attribute: 'ref', reason: 'the used range' }],
])

/** What a deletion would destroy that cannot survive as #REF!, or undefined. */
export function deletionDamage(bytes: Uint8Array, spec: ShiftSpec): string | undefined {
  let row = 0
  let column = 0
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'row') {
      const declared = event.attributes.get('r')
      row = declared === undefined ? row + 1 : Number(declared)
      column = 0
      continue
    }
    if (event.localName === 'c') {
      const letters = event.attributes.get('r')?.match(/[A-Za-z]+/)?.[0]
      column = letters === undefined ? column + 1 : columnToIndex(letters)
    }
    // A shared or array formula's home cell inside the band loses its definition.
    if (event.localName === 'f' && event.attributes.get('ref') !== undefined) {
      const line = spec.axis === 'row' ? row : column
      if (line >= spec.at && line < spec.at - spec.delta) return 'a shared or array formula'
    }
    const collapsible = COLLAPSIBLE.get(event.localName)
    if (collapsible !== undefined) {
      const value = event.attributes.get(collapsible.attribute)
      if (value !== undefined && shiftFormula(value, spec).includes('#REF!'))
        return collapsible.reason
    }
  }
  return undefined
}

// Parts whose stored positions a line shift does not adjust, so an insert or
// delete that would move the cells under them is refused rather than leave them
// pointing at the wrong ones. A drawing pins charts, images and shapes by cell
// anchor; a comment and a table each pin a range; a pivot pins its source and
// where it lands. The capital T in pivotTable keeps it from matching table.
const UNSHIFTABLE_PARTS: ReadonlyArray<readonly [string, string]> = [
  ['relationships/table', 'a table'],
  ['relationships/pivotTable', 'a pivot table'],
  ['relationships/drawing', 'a drawing'],
  ['relationships/vmlDrawing', 'a drawing'],
  ['relationships/comments', 'a comment'],
]

/** The kind of unshiftable part a sheet owns, named for the refusal, or undefined. */
export function unshiftablePart(relationshipsXml: string): string | undefined {
  for (const event of readXml(relationshipsXml)) {
    if (event.kind !== 'open' || event.localName !== 'Relationship') continue
    for (const [suffix, noun] of UNSHIFTABLE_PARTS) {
      if (event.attributes.get('Type')?.endsWith(suffix) === true) return noun
    }
  }
  return undefined
}
