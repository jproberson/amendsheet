import { readXml, readXmlBytes } from './xml.js'

/** The page setup for printing: which way the paper turns and the print scale. */
export interface PageSetup {
  readonly orientation?: 'portrait' | 'landscape'
  /** Print scale as a whole percentage, 10 to 400. */
  readonly scale?: number
}

/** The print margins, in inches. Every field is optional. */
export interface PageMargins {
  readonly left?: number
  readonly right?: number
  readonly top?: number
  readonly bottom?: number
  readonly header?: number
  readonly footer?: number
}

const MARGIN_KEYS = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const
// Excel's default margins, so a partial edit still writes a complete element.
const DEFAULT_MARGINS: Required<PageMargins> = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
}

// The worksheet children that follow pageMargins, in schema order, so a fresh
// element lands before the first one present and after everything before it.
const AFTER_PAGE_MARGINS =
  'pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF drawingHF picture oleObjects controls webPublishItems tableParts extLst'.split(
    ' ',
  )

const numberAttribute = (
  attributes: ReadonlyMap<string, string>,
  name: string,
): number | undefined => {
  const raw = attributes.get(name)
  if (raw === undefined) return undefined
  const value = Number(raw)
  return Number.isFinite(value) ? value : undefined
}

/** Sets or replaces an attribute on an open tag, keeping the rest. */
function setAttribute(openTag: string, name: string, value: string | number): string {
  const pattern = new RegExp(`(\\s${name}\\s*=\\s*)("[^"]*"|'[^']*')`)
  if (pattern.test(openTag)) return openTag.replace(pattern, `$1"${value}"`)
  return openTag.replace(/^<([^\s/>]+)/, `<$1 ${name}="${value}"`)
}

/** Inserts a fresh self-closing element before the first of `successors` present,
 * or before the worksheet's close when it has none. */
function insertBeforeSuccessor(
  sheetXml: string,
  element: string,
  successors: readonly string[],
): string {
  let at = sheetXml.indexOf('</worksheet')
  if (at === -1) at = sheetXml.length
  for (const successor of successors) {
    const found = sheetXml.search(new RegExp(`<(?:\\w+:)?${successor}\\b`))
    if (found !== -1 && found < at) at = found
  }
  return `${sheetXml.slice(0, at)}${element}${sheetXml.slice(at)}`
}

function selfClosingTag(
  sheetXml: string,
  localName: string,
): { start: number; end: number } | undefined {
  const match = new RegExp(`<(?:\\w+:)?${localName}\\b[^>]*/>`).exec(sheetXml)
  if (match === null) return undefined
  return { start: match.index, end: match.index + match[0].length }
}

/**
 * Writes the page margins, merging onto the ones the sheet has (or Excel's
 * defaults), so a partial edit still leaves a complete element. All six values
 * are written, in inches.
 */
export function withPageMargins(sheetXml: string, margins: PageMargins): string {
  const tag = selfClosingTag(sheetXml, 'pageMargins')
  const current: { -readonly [K in keyof Required<PageMargins>]: number } = { ...DEFAULT_MARGINS }
  if (tag !== undefined) {
    for (const event of readXml(sheetXml)) {
      if (event.kind === 'open' && event.localName === 'pageMargins') {
        for (const key of MARGIN_KEYS) {
          const value = numberAttribute(event.attributes, key)
          if (value !== undefined) current[key] = value
        }
      }
    }
  }
  for (const key of MARGIN_KEYS) {
    const value = margins[key]
    if (value !== undefined) current[key] = value
  }
  let attributes = ''
  for (const key of MARGIN_KEYS) attributes += ` ${key}="${current[key]}"`
  const element = `<pageMargins${attributes}/>`
  if (tag !== undefined) return sheetXml.slice(0, tag.start) + element + sheetXml.slice(tag.end)
  return insertBeforeSuccessor(sheetXml, element, AFTER_PAGE_MARGINS)
}

/**
 * Writes the page setup, setting only the orientation and scale a caller gives and
 * keeping any other attribute the sheet's `pageSetup` already carries.
 */
export function withPageSetup(sheetXml: string, setup: PageSetup): string {
  const tag = selfClosingTag(sheetXml, 'pageSetup')
  if (tag !== undefined) {
    let openTag = sheetXml.slice(tag.start, tag.end)
    if (setup.orientation !== undefined)
      openTag = setAttribute(openTag, 'orientation', setup.orientation)
    if (setup.scale !== undefined) openTag = setAttribute(openTag, 'scale', setup.scale)
    return sheetXml.slice(0, tag.start) + openTag + sheetXml.slice(tag.end)
  }
  const orientation = setup.orientation === undefined ? '' : ` orientation="${setup.orientation}"`
  const scale = setup.scale === undefined ? '' : ` scale="${setup.scale}"`
  return insertBeforeSuccessor(
    sheetXml,
    `<pageSetup${orientation}${scale}/>`,
    AFTER_PAGE_MARGINS.slice(1),
  )
}

/** Reads the page setup — the orientation and scale this models. */
export function readPageSetup(bytes: Uint8Array): PageSetup {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'pageSetup') continue
    const orientation = event.attributes.get('orientation')
    const scale = numberAttribute(event.attributes, 'scale')
    const result: { -readonly [K in keyof PageSetup]?: PageSetup[K] } = {}
    if (orientation === 'portrait' || orientation === 'landscape') result.orientation = orientation
    if (scale !== undefined) result.scale = scale
    return result
  }
  return {}
}

/** Reads the page margins the sheet stores. */
export function readPageMargins(bytes: Uint8Array): PageMargins {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'pageMargins') continue
    const margins: { -readonly [K in keyof PageMargins]?: number } = {}
    for (const key of MARGIN_KEYS) {
      const value = numberAttribute(event.attributes, key)
      if (value !== undefined) margins[key] = value
    }
    return margins
  }
  return {}
}
