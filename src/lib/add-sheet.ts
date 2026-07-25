import { XlsxError } from './errors.js'
import type { SheetRef } from './workbook.js'

const WORKSHEET_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'

/** An empty worksheet part, the base a freshly added sheet's edits are applied to. */
export const EMPTY_SHEET_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'

/** A sheet added since the read, with the wiring toBytes() needs to attach it. */
export interface AddedSheet {
  readonly reference: SheetRef
  /** The r:id the workbook's `<sheet>` and its relationship share. */
  readonly relationshipId: string
  /** The relationship target: the sheet part relative to the workbook's folder. */
  readonly target: string
}

/**
 * Refuses a sheet name Excel will not take, at the `addSheet()` call. Types stop
 * a TypeScript caller passing a non-string; a JS caller or an `any` reaches here.
 */
export function checkSheetName(name: unknown, existing: readonly string[]): void {
  if (typeof name !== 'string' || name.length === 0) {
    throw new XlsxError('unwritable-value', 'A sheet name must be a non-empty string', {})
  }
  if (name.length > 31) {
    throw new XlsxError('unwritable-value', `Sheet name "${name}" is longer than 31 characters`, {})
  }
  if (/[:\\/?*[\]]/.test(name)) {
    throw new XlsxError(
      'unwritable-value',
      `Sheet name "${name}" holds a character Excel forbids in a sheet name (: \\ / ? * [ ])`,
      {},
    )
  }
  const taken = name.toLowerCase()
  if (existing.some((other) => other.toLowerCase() === taken)) {
    throw new XlsxError('unwritable-value', `A sheet named "${name}" already exists`, {})
  }
}

// A sheet name is written into an attribute, so the markup characters are
// escaped; the caller-facing addSheet() has already refused the ones Excel bans.
export const escapeSheetName = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const sheetElement = (prefix: string, sheet: AddedSheet) =>
  `<${prefix}sheet name="${escapeSheetName(sheet.reference.name)}" sheetId="${sheet.reference.sheetId}"` +
  ` r:id="${sheet.relationshipId}"/>`

/** Adds each new `<sheet>` to the workbook's `<sheets>`, opening a self-closing one. */
export function withSheetsAdded(workbookXml: string, added: readonly AddedSheet[]): string {
  if (added.length === 0) return workbookXml

  const selfClosing = workbookXml.match(/<([a-z0-9]*:)?sheets\s*\/>/i)
  if (selfClosing !== null) {
    const prefix = selfClosing[1] ?? ''
    const inner = added.map((sheet) => sheetElement(prefix, sheet)).join('')
    return workbookXml.replace(selfClosing[0], `<${prefix}sheets>${inner}</${prefix}sheets>`)
  }

  const close = workbookXml.match(/<\/([a-z0-9]*:)?sheets>/i)
  if (close === null) {
    throw new XlsxError('invalid-content', 'Workbook has no sheets element to add a sheet to', {
      part: 'xl/workbook.xml',
    })
  }
  const prefix = close[1] ?? ''
  const inner = added.map((sheet) => sheetElement(prefix, sheet)).join('')
  return workbookXml.replace(close[0], `${inner}${close[0]}`)
}

const escapeRegExp = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** Renames the `<sheet>` carrying `oldName` in the workbook's `<sheets>`. */
export function withSheetRenamed(workbookXml: string, oldName: string, newName: string): string {
  const pattern = new RegExp(
    `(<([a-z0-9]*:)?sheet\\b[^>]*\\bname=")${escapeRegExp(escapeSheetName(oldName))}"`,
    'i',
  )
  return workbookXml.replace(pattern, (_match, lead) => `${lead}${escapeSheetName(newName)}"`)
}

/** Adds a relationship for each new sheet to the workbook's relationships part. */
export function withSheetRelationships(relsXml: string, added: readonly AddedSheet[]): string {
  if (added.length === 0) return relsXml
  const inner = added
    .map(
      (sheet) =>
        `<Relationship Id="${sheet.relationshipId}" Type="${WORKSHEET_RELATIONSHIP}"` +
        ` Target="${sheet.target}"/>`,
    )
    .join('')
  const close = relsXml.match(/<\/Relationships>/)
  if (close === null) {
    throw new XlsxError('invalid-content', 'Workbook relationships part is malformed', {})
  }
  return relsXml.replace(close[0], `${inner}${close[0]}`)
}

/** Declares each new sheet part in `[Content_Types].xml` with an Override. */
export function withSheetContentTypes(
  contentTypesXml: string,
  added: readonly AddedSheet[],
): string {
  if (added.length === 0) return contentTypesXml
  const inner = added
    .map(
      (sheet) =>
        `<Override PartName="/${sheet.reference.path}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`,
    )
    .join('')
  const close = contentTypesXml.match(/<\/Types>/)
  if (close === null) {
    throw new XlsxError('invalid-content', 'Content types part is malformed', {
      part: '[Content_Types].xml',
    })
  }
  return contentTypesXml.replace(close[0], `${inner}${close[0]}`)
}
