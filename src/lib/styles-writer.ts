import { XlsxError } from './errors.js'
import { builtInFormatId, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { findUnwritableCharacter, readXml, withAttribute } from './xml.js'

/** Built in short date format, so no custom numFmt has to be written. */
const SHORT_DATE_FORMAT_ID = 14

/** Custom format ids start here; everything below is built in. */
const FIRST_CUSTOM_FORMAT_ID = 164

const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export interface DateStyle {
  readonly xml: string
  /** Index a cell should carry in its `s` attribute to display as a date. */
  readonly index: number
}

const prefixOf = (name: string) => {
  const colon = name.indexOf(':')
  return colon === -1 ? '' : name.slice(0, colon + 1)
}

/** One of the styles sub-tables — `fonts`, `fills`, `borders`, `cellXfs`. */
interface StyleTable {
  readonly elements: string[]
  /** Namespace prefix the file writes its elements with, `x:` or empty. */
  readonly prefix: string
  readonly openTag: string
  readonly openStart: number
  readonly openEnd: number
  readonly insertAt: number
  readonly selfClosing: boolean
}

/** Reads the direct `child` elements of `container`, each kept whole so a clone
 * carries its own children (an xf's alignment, a font's parts) with it. */
function readTable(xml: string, container: string, child: string): StyleTable | undefined {
  let openTag = ''
  let openStart = -1
  let openEnd = -1
  let insertAt = -1
  let selfClosing = false
  let inside = false
  let prefix = ''
  let openElement = -1
  const elements: string[] = []

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === container) {
      prefix = prefixOf(event.name)
      openTag = xml.slice(event.start, event.end)
      openStart = event.start
      openEnd = event.end
      selfClosing = event.selfClosing
      inside = !event.selfClosing
      if (selfClosing) insertAt = event.end
      continue
    }
    if (event.kind === 'close' && event.localName === container) {
      insertAt = event.start
      inside = false
      continue
    }
    if (inside && event.kind === 'open' && event.localName === child) {
      if (event.selfClosing) elements.push(xml.slice(event.start, event.end))
      else openElement = event.start
      continue
    }
    if (inside && event.kind === 'close' && event.localName === child && openElement !== -1) {
      elements.push(xml.slice(openElement, event.end))
      openElement = -1
    }
  }

  if (openStart === -1) return undefined
  return { elements, prefix, openTag, openStart, openEnd, insertAt, selfClosing }
}

/** Where a new sub-table is opened: before `cellXfs` so fonts/fills/borders keep
 * their place ahead of it, or before the closing tag when there is no cellXfs. */
function tableInsertPoint(xml: string): { position: number; prefix: string } {
  let cellXfsStart = -1
  let closeStart = -1
  let prefix = ''
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'cellXfs' && cellXfsStart === -1) {
      cellXfsStart = event.start
    }
    if (event.kind === 'close' && event.localName === 'styleSheet') {
      closeStart = event.start
      prefix = prefixOf(event.name)
    }
  }
  if (closeStart === -1) {
    throw new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
      part: 'xl/styles.xml',
    })
  }
  return { position: cellXfsStart === -1 ? closeStart : cellXfsStart, prefix }
}

/** Adds `element` to `container`, or returns the index of an identical one. */
function ensureInTable(
  xml: string,
  container: string,
  child: string,
  element: string,
): { xml: string; id: number } {
  const table = readTable(xml, container, child)
  if (table === undefined) {
    const { position, prefix } = tableInsertPoint(xml)
    const prefixed = element.replace(/^<[^\s/>]+/, `<${prefix}${child}`)
    const created = `<${prefix}${container} count="1">${prefixed}</${prefix}${container}>`
    return { xml: xml.slice(0, position) + created + xml.slice(position), id: 0 }
  }

  const prefixed = element.replace(/^<[^\s/>]+/, `<${table.prefix}${child}`)
  const existing = table.elements.indexOf(prefixed)
  if (existing !== -1) return { xml, id: existing }

  const id = table.elements.length
  const openTag = withAttribute(table.openTag, 'count', id + 1)

  if (table.selfClosing) {
    const opened = `${openTag.slice(0, -2)}>`
    return {
      xml:
        xml.slice(0, table.openStart) +
        `${opened}${prefixed}</${table.prefix}${container}>` +
        xml.slice(table.openEnd),
      id,
    }
  }

  const head = xml.slice(0, table.openStart) + openTag
  const body = xml.slice(table.openEnd, table.insertAt)
  return { xml: `${head}${body}${prefixed}${xml.slice(table.insertAt)}`, id }
}

const DEFAULT_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

/** The parts of a cell format this library sets, each with its `apply*` flag. */
interface CellOverrides {
  readonly numFmtId?: number
  readonly fontId?: number
  readonly fillId?: number
  readonly borderId?: number
}

const OVERRIDE_FLAGS: ReadonlyArray<readonly [keyof CellOverrides, string]> = [
  ['numFmtId', 'applyNumberFormat'],
  ['fontId', 'applyFont'],
  ['fillId', 'applyFill'],
  ['borderId', 'applyBorder'],
]

/** Turns an `apply*` flag on, whether it was off, absent, or written as a word. */
function withApplyFlag(openTag: string, flag: string): string {
  const pattern = new RegExp(`${flag}=("|')(?:\\d+|true|false)\\1`)
  if (pattern.test(openTag)) return openTag.replace(pattern, `${flag}="1"`)
  return openTag.replace(/\/?>$/, (tag) => (tag === '/>' ? ` ${flag}="1"/>` : ` ${flag}="1">`))
}

function withOverrides(element: string, overrides: CellOverrides): string {
  const openTagEnd = element.indexOf('>') + 1
  let openTag = element.slice(0, openTagEnd)
  const rest = element.slice(openTagEnd)

  for (const [attribute, flag] of OVERRIDE_FLAGS) {
    const value = overrides[attribute]
    // The element may be written <xf> or <x:xf>, so the name is matched rather
    // than assumed. Missing it left the cell on General while it held a serial.
    if (value !== undefined) openTag = withApplyFlag(withAttribute(openTag, attribute, value), flag)
  }

  return `${openTag}${rest}`
}

/**
 * Finds a cell format that displays dates, adding one if the file has none.
 * `basedOn` keeps the font, fill and border of the cell being written to, so
 * applying a date format does not strip the rest of its formatting.
 */
export function ensureDateStyle(stylesXml: string, basedOn: number | undefined): DateStyle {
  const parsed = readStyles(stylesXml)
  if (basedOn !== undefined && isDateFormat(parsed, basedOn)) {
    return { xml: stylesXml, index: basedOn }
  }

  // A cell with no formatting of its own can borrow any date format already here.
  if (basedOn === undefined) {
    for (let index = 0; index < parsed.cellFormats.length; index++) {
      if (isDateFormat(parsed, index)) return { xml: stylesXml, index }
    }
  }

  return applyCellFormat(stylesXml, basedOn, { numFmtId: SHORT_DATE_FORMAT_ID })
}

/** Puts a cell format carrying `overrides` in the table, reusing one if it fits. */
function applyCellFormat(
  stylesXml: string,
  basedOn: number | undefined,
  overrides: CellOverrides,
): DateStyle {
  const formats = readTable(stylesXml, 'cellXfs', 'xf')
  const base = basedOn === undefined ? DEFAULT_XF : (formats?.elements[basedOn] ?? DEFAULT_XF)
  const wanted = withOverrides(base, overrides)
  const { xml, id } = ensureInTable(stylesXml, 'cellXfs', 'xf', wanted)
  return { xml, index: id }
}

export interface FontFormat {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly size?: number
  /** An `RRGGBB` or `AARRGGBB` hex colour; a six-digit value is stored opaque. */
  readonly color?: string
  readonly name?: string
}

/** ECMA-376 stores a colour as eight hex digits, alpha first; a six-digit value
 * is the colour at full opacity. */
export function normalizeColor(color: string): string {
  const hex = color.startsWith('#') ? color.slice(1) : color
  if (!/^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(hex)) {
    throw new XlsxError('unwritable-value', `Colour "${color}" is not a 6 or 8 digit hex value`, {})
  }
  return (hex.length === 6 ? `FF${hex}` : hex).toUpperCase()
}

/** A `<b val="0"/>` turns bold off, so presence alone is not the whole story. */
const flagOn = (value: string | undefined) => value !== '0' && value !== 'false'

function parseFont(element: string): FontFormat {
  const font: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    size?: number
    color?: string
    name?: string
  } = {}
  for (const event of readXml(element)) {
    if (event.kind !== 'open') continue
    switch (event.localName) {
      case 'b':
        font.bold = flagOn(event.attributes.get('val'))
        break
      case 'i':
        font.italic = flagOn(event.attributes.get('val'))
        break
      case 'u':
        font.underline = flagOn(event.attributes.get('val'))
        break
      case 'sz': {
        const size = Number(event.attributes.get('val'))
        if (!Number.isNaN(size)) font.size = size
        break
      }
      case 'color': {
        const rgb = event.attributes.get('rgb')
        if (rgb !== undefined) font.color = rgb
        break
      }
      case 'name': {
        const name = event.attributes.get('val')
        if (name !== undefined) font.name = name
        break
      }
    }
  }
  return font
}

function buildFontElement(font: FontFormat): string {
  let inner = ''
  if (font.bold === true) inner += '<b/>'
  if (font.italic === true) inner += '<i/>'
  if (font.underline === true) inner += '<u/>'
  if (font.size !== undefined) inner += `<sz val="${font.size}"/>`
  if (font.color !== undefined) inner += `<color rgb="${normalizeColor(font.color)}"/>`
  if (font.name !== undefined) inner += `<name val="${escapeXml(font.name)}"/>`
  return `<font>${inner}</font>`
}

/** The font id a cell format points at, or 0 (the default font) when it names none. */
function fontIdOf(stylesXml: string, basedOn: number | undefined): number {
  if (basedOn === undefined) return 0
  const element = readTable(stylesXml, 'cellXfs', 'xf')?.elements[basedOn]
  const match = element?.match(/\bfontId\s*=\s*["'](\d+)["']/)
  return match?.[1] === undefined ? 0 : Number(match[1])
}

/**
 * Applies `font` to a cell, merging onto the font it already has so setting bold
 * does not reset its size or colour. The merged font is added if the file has no
 * identical one, and a cell format pointing at it is returned.
 */
export function ensureFontStyle(
  stylesXml: string,
  basedOn: number | undefined,
  font: FontFormat,
): DateStyle {
  const current = parseFont(
    readTable(stylesXml, 'fonts', 'font')?.elements[fontIdOf(stylesXml, basedOn)] ?? '',
  )
  const merged: FontFormat = {
    bold: font.bold ?? current.bold,
    italic: font.italic ?? current.italic,
    underline: font.underline ?? current.underline,
    size: font.size ?? current.size,
    color: font.color ?? current.color,
    name: font.name ?? current.name,
  }

  const { xml, id } = ensureInTable(stylesXml, 'fonts', 'font', buildFontElement(merged))
  return applyCellFormat(xml, basedOn, { fontId: id })
}

export interface FillFormat {
  /** Solid fill colour, `RRGGBB` or `AARRGGBB` hex. */
  readonly color: string
}

// fillId 0 (none) and 1 (gray125) are reserved, so a real solid fill is the
// third entry; a file with no fills table has these seeded before ours is added.
const RESERVED_FILLS =
  '<fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>'

function withReservedFills(stylesXml: string): string {
  if (readTable(stylesXml, 'fills', 'fill') !== undefined) return stylesXml
  const { position, prefix } = tableInsertPoint(stylesXml)
  const seeded = RESERVED_FILLS.replace(/<(\/?)(fill|patternFill)/g, `<$1${prefix}$2`)
  return `${stylesXml.slice(0, position)}<${prefix}fills count="2">${seeded}</${prefix}fills>${stylesXml.slice(position)}`
}

export function ensureFillStyle(
  stylesXml: string,
  basedOn: number | undefined,
  fill: FillFormat,
): DateStyle {
  const seeded = withReservedFills(stylesXml)
  const element = `<fill><patternFill patternType="solid"><fgColor rgb="${normalizeColor(fill.color)}"/><bgColor indexed="64"/></patternFill></fill>`
  const { xml, id } = ensureInTable(seeded, 'fills', 'fill', element)
  return applyCellFormat(xml, basedOn, { fillId: id })
}

function tablePrefix(xml: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName !== 'numFmts' && event.localName !== 'styleSheet') continue
    const colon = event.name.indexOf(':')
    return colon === -1 ? '' : event.name.slice(0, colon + 1)
  }
  return ''
}

/** Every numFmtId the file mentions, including ones only used by dxfs. */
function usedFormatIds(xml: string): Set<number> {
  const used = new Set<number>()
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    const declared = event.attributes.get('numFmtId')
    if (declared === undefined) continue
    const id = Number(declared)
    if (!Number.isNaN(id)) used.add(id)
  }
  return used
}

/** Adds a numFmt element, creating the numFmts table when the file has none. */
function withNumberFormat(xml: string, id: number, code: string, prefix: string): string {
  const element = `<${prefix}numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`
  let openStart = -1
  let openEnd = -1
  // Counted rather than read off the open tag, because a file that omits count
  // or gets it wrong would otherwise have that error carried forward.
  let children = 0

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'numFmts' && event.selfClosing) {
      const opened = `${withAttribute(xml.slice(event.start, event.end - 2), 'count', 1)}>`
      return `${xml.slice(0, event.start)}${opened}${element}</${prefix}numFmts>${xml.slice(event.end)}`
    }
    if (event.kind === 'open' && event.localName === 'numFmts') {
      openStart = event.start
      openEnd = event.end
      continue
    }
    if (openStart !== -1 && event.kind === 'open' && event.localName === 'numFmt') {
      children++
      continue
    }
    // openStart guards a close with no open, which leaves the table to be
    // created below rather than splicing against a position that was never set.
    if (event.kind === 'close' && event.localName === 'numFmts' && openStart !== -1) {
      // A count that disagrees with the children makes Excel call the file
      // unreadable and offer to repair it, which rewrites the whole package.
      const head =
        xml.slice(0, openStart) +
        withAttribute(xml.slice(openStart, openEnd), 'count', children + 1) +
        xml.slice(openEnd, event.start)
      return head + element + xml.slice(event.start)
    }
  }

  // No table at all, so one is opened directly after the root element.
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'styleSheet') continue
    const table = `<${prefix}numFmts count="1">${element}</${prefix}numFmts>`
    return xml.slice(0, event.end) + table + xml.slice(event.end)
  }

  throw new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
    part: 'xl/styles.xml',
  })
}

/**
 * Finds a cell format that shows `formatCode`, adding the format and the cell
 * format when the file has neither. `basedOn` keeps the font, fill and border
 * of the cell being written to.
 */
export function ensureNumberFormat(
  stylesXml: string,
  basedOn: number | undefined,
  formatCode: string,
): DateStyle {
  const unwritable = findUnwritableCharacter(formatCode)
  if (unwritable !== undefined) {
    throw new XlsxError(
      'unwritable-value',
      `Number format holds ${unwritable}, which cannot be written to xml`,
      { part: 'xl/styles.xml' },
    )
  }

  const parsed = readStyles(stylesXml)
  if (basedOn !== undefined && numberFormatOf(parsed, basedOn) === formatCode) {
    return { xml: stylesXml, index: basedOn }
  }
  // Only a cell with no formatting of its own can borrow another cell's format;
  // otherwise the borrowed xf brings its font, fill and border along and the
  // target's are silently dropped.
  if (basedOn === undefined) {
    for (let index = 0; index < parsed.cellFormats.length; index++) {
      if (numberFormatOf(parsed, index) === formatCode) return { xml: stylesXml, index }
    }
  }

  let formatId = builtInFormatId(formatCode)
  let withFormat = stylesXml
  const prefix = tablePrefix(stylesXml)

  if (formatId === undefined) {
    for (const [id, code] of parsed.numberFormats) {
      if (code === formatCode) formatId = id
    }
  }
  if (formatId === undefined) {
    const used = usedFormatIds(stylesXml)
    formatId = FIRST_CUSTOM_FORMAT_ID
    while (used.has(formatId)) formatId++
    withFormat = withNumberFormat(stylesXml, formatId, formatCode, prefix)
  }

  return applyCellFormat(withFormat, basedOn, { numFmtId: formatId })
}
