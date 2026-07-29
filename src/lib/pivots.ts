import { type ShiftSpec, shiftFormula } from './shift.js'
import { readXml } from './xml.js'

const setRef = (tag: string, reference: string): string =>
  tag.replace(
    /(\sref\s*=\s*)("[^"]*"|'[^']*')/,
    (_, head: string, quoted: string) => `${head}${quoted[0]}${reference}${quoted[0]}`,
  )

// Rewrites the `ref` on the first element with the given name, shifting the range
// it holds as a reference on the edited sheet. Elements without a `ref` are left be.
const shiftRefOf = (xml: string, localName: string, spec: ShiftSpec): string => {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== localName) continue
    const reference = event.attributes.get('ref')
    if (reference === undefined) return xml
    const shifted = shiftFormula(reference, { ...spec, onCurrentSheet: true })
    if (shifted === reference) return xml
    const tag = setRef(xml.slice(event.start, event.end), shifted)
    return xml.slice(0, event.start) + tag + xml.slice(event.end)
  }
  return xml
}

/**
 * Moves a pivot table's `location` — the range on its own sheet where it is drawn
 * — with the rows or columns an edit shifts under it. Called only for a pivot table
 * on the edited sheet, so its location is a reference on that sheet.
 */
export function shiftPivotLocation(pivotXml: string, spec: ShiftSpec): string {
  return shiftRefOf(pivotXml, 'location', spec)
}

/**
 * Moves a pivot cache's `worksheetSource` range when the sheet it reads is the one
 * an edit changed, so a later refresh reads the data where it now sits. A source
 * given by name — a table or defined name — carries no `ref` and moves with that
 * name instead. The sheet is matched case-insensitively, the way Excel names it.
 */
export function shiftPivotCacheSource(cacheXml: string, spec: ShiftSpec): string {
  for (const event of readXml(cacheXml)) {
    if (event.kind !== 'open' || event.localName !== 'worksheetSource') continue
    const sheet = event.attributes.get('sheet')
    if (sheet === undefined || sheet.toLowerCase() !== spec.editedSheet.toLowerCase())
      return cacheXml
    return shiftRefOf(cacheXml, 'worksheetSource', spec)
  }
  return cacheXml
}
