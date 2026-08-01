import { escapeSheetName } from './add-sheet.js'
import { XlsxError } from './errors.js'
import { readXml } from './xml.js'

/**
 * Refuses a defined name Excel will not take, at the `defineName()` call. A name
 * starts with a letter, an underscore or a backslash and then holds only letters,
 * digits, periods, underscores or backslashes — no spaces, up to 255 characters.
 */
export function checkDefinedName(name: unknown, refersTo: unknown): void {
  if (
    typeof name !== 'string' ||
    !/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name) ||
    name.length > 255
  ) {
    throw new XlsxError('unwritable-value', `"${String(name)}" is not a name Excel accepts`, {})
  }
  if (typeof refersTo !== 'string' || refersTo.length === 0) {
    throw new XlsxError('unwritable-value', `A defined name must refer to a non-empty formula`, {})
  }
}

/**
 * The workbook's global defined names, each mapped to what it refers to. A name
 * scoped to one sheet (it carries a `localSheetId`) is left out, since two sheets
 * can hold the same one.
 */
// A name scoped to one sheet carries a localSheetId; only a global one is ours.
const globalName = (attributes: ReadonlyMap<string, string>): string | undefined => {
  const name = attributes.get('name')
  return name !== undefined && attributes.get('localSheetId') === undefined ? name : undefined
}

/** The reserved prefix Excel gives built-in names — print area, print titles,
 * filter database — which carry their own accessors, not `defineName`. */
export const isBuiltInName = (name: string): boolean => name.startsWith('_xlnm.')

export interface SheetScopedName {
  readonly name: string
  /** The 0-based index into the workbook's sheet order the name is scoped to. */
  readonly localSheetId: number
  readonly refersTo: string
}

/**
 * Every defined name carrying a `localSheetId`, in file order, the built-in
 * `_xlnm.*` names (print area, print titles) included — a caller filters. The
 * global names, which `readDefinedNames` returns, are skipped here.
 */
export function readSheetScopedNames(workbookXml: string): SheetScopedName[] {
  const names: SheetScopedName[] = []
  let current: string | undefined
  let localSheetId = -1
  let refersTo = ''
  for (const event of readXml(workbookXml)) {
    if (event.kind === 'open' && event.localName === 'definedName') {
      const name = event.attributes.get('name')
      const scoped = event.attributes.get('localSheetId')
      const id = scoped === undefined ? Number.NaN : Number(scoped)
      current = name !== undefined && Number.isInteger(id) && id >= 0 ? name : undefined
      localSheetId = id
      refersTo = ''
      continue
    }
    if (event.kind === 'text' && current !== undefined) {
      refersTo += event.text
      continue
    }
    if (event.kind === 'close' && event.localName === 'definedName' && current !== undefined) {
      names.push({ name: current, localSheetId, refersTo })
      current = undefined
    }
  }
  return names
}

export function readDefinedNames(workbookXml: string): Map<string, string> {
  const names = new Map<string, string>()
  let current: string | undefined
  let refersTo = ''
  for (const event of readXml(workbookXml)) {
    if (event.kind === 'open' && event.localName === 'definedName') {
      current = globalName(event.attributes)
      refersTo = ''
      continue
    }
    if (event.kind === 'text' && current !== undefined) {
      refersTo += event.text
      continue
    }
    if (event.kind === 'close' && event.localName === 'definedName' && current !== undefined) {
      names.set(current, refersTo)
      current = undefined
    }
  }
  return names
}

/** Adds each global defined name to the workbook, replacing one of the same name. */
export function withDefinedNames(
  workbookXml: string,
  names: ReadonlyMap<string, string>,
  removed: ReadonlySet<string> = new Set(),
): string {
  if (names.size === 0 && removed.size === 0) return workbookXml

  // Drop any existing global definedName we are redefining (so it is not doubled)
  // or removing (so it is gone). Only the redefined ones are written back below.
  const spans: { start: number; end: number }[] = []
  let redefined: number | undefined
  for (const event of readXml(workbookXml)) {
    if (event.kind === 'open' && event.localName === 'definedName') {
      const name = globalName(event.attributes)
      redefined =
        name !== undefined && (names.has(name) || removed.has(name)) ? event.start : undefined
      continue
    }
    if (event.kind === 'close' && event.localName === 'definedName' && redefined !== undefined) {
      spans.push({ start: redefined, end: event.end })
      redefined = undefined
    }
  }
  let xml = workbookXml
  for (const span of spans.sort((a, b) => b.start - a.start)) {
    xml = xml.slice(0, span.start) + xml.slice(span.end)
  }

  const elements = (prefix: string) =>
    [...names]
      .map(
        ([name, refersTo]) =>
          `<${prefix}definedName name="${escapeSheetName(name)}">${escapeSheetName(refersTo)}</${prefix}definedName>`,
      )
      .join('')

  const close = xml.match(/<\/([a-z0-9]*:)?definedNames>/i)
  if (close !== null) return xml.replace(close[0], `${elements(close[1] ?? '')}${close[0]}`)

  const selfClosing = xml.match(/<([a-z0-9]*:)?definedNames\s*\/>/i)
  if (selfClosing !== null) {
    const prefix = selfClosing[1] ?? ''
    return xml.replace(
      selfClosing[0],
      `<${prefix}definedNames>${elements(prefix)}</${prefix}definedNames>`,
    )
  }

  // definedNames sits after sheets, so a fresh one opens right after </sheets>.
  const sheets = xml.match(/<\/([a-z0-9]*:)?sheets>/i)
  if (sheets === null) {
    throw new XlsxError('invalid-content', 'Workbook has no sheets element to place names after', {
      part: 'xl/workbook.xml',
    })
  }
  const prefix = sheets[1] ?? ''
  return xml.replace(
    sheets[0],
    `${sheets[0]}<${prefix}definedNames>${elements(prefix)}</${prefix}definedNames>`,
  )
}
