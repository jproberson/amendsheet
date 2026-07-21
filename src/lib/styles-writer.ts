import { XlsxError } from './errors.js'
import { builtInFormatId, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { findUnwritableCharacter, readXml } from './xml.js'

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

interface CellFormats {
  readonly elements: string[]
  /** Namespace prefix the file writes its elements with, `x:` or empty. */
  readonly prefix: string
  readonly openTag: string
  readonly openStart: number
  readonly openEnd: number
  readonly insertAt: number
  readonly selfClosing: boolean
}

function readCellFormats(xml: string): CellFormats | undefined {
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
    if (event.kind === 'open' && event.localName === 'cellXfs') {
      const colon = event.name.indexOf(':')
      prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      openTag = xml.slice(event.start, event.end)
      openStart = event.start
      openEnd = event.end
      selfClosing = event.selfClosing
      inside = !event.selfClosing
      if (selfClosing) insertAt = event.end
      continue
    }
    if (event.kind === 'close' && event.localName === 'cellXfs') {
      insertAt = event.start
      inside = false
      continue
    }
    if (inside && event.kind === 'open' && event.localName === 'xf') {
      if (event.selfClosing) elements.push(xml.slice(event.start, event.end))
      else openElement = event.start
      continue
    }
    // A style may carry alignment or protection children, which have to come
    // along or the cloned element is an unclosed tag.
    if (inside && event.kind === 'close' && event.localName === 'xf' && openElement !== -1) {
      elements.push(xml.slice(openElement, event.end))
      openElement = -1
    }
  }

  if (openStart === -1) return undefined
  return { elements, prefix, openTag, openStart, openEnd, insertAt, selfClosing }
}

function withFormatId(element: string, formatId: number): string {
  const openTagEnd = element.indexOf('>') + 1
  const openTag = element.slice(0, openTagEnd)
  const rest = element.slice(openTagEnd)

  const withFormat = openTag.includes('numFmtId="')
    ? openTag.replace(/numFmtId="\d+"/, `numFmtId="${formatId}"`)
    : openTag.replace(/^<xf/, `<xf numFmtId="${formatId}"`)

  const applied = withFormat.includes('applyNumberFormat=')
    ? withFormat.replace(/applyNumberFormat="(?:\d+|true|false)"/, 'applyNumberFormat="1"')
    : withFormat.replace(/\/?>$/, (tag) =>
        tag === '/>' ? ' applyNumberFormat="1"/>' : ' applyNumberFormat="1">',
      )

  return `${applied}${rest}`
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

  return applyFormatId(stylesXml, basedOn, SHORT_DATE_FORMAT_ID)
}

/** Puts a cell format carrying `formatId` in the table, reusing one if it fits. */
function applyFormatId(
  stylesXml: string,
  basedOn: number | undefined,
  formatId: number,
): DateStyle {
  const formats = readCellFormats(stylesXml)
  const defaultXf = `<xf numFmtId="${formatId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`
  const wanted =
    basedOn === undefined
      ? defaultXf
      : withFormatId(formats?.elements[basedOn] ?? defaultXf, formatId)

  if (formats === undefined) {
    let closeStart = -1
    let rootPrefix = ''
    for (const event of readXml(stylesXml)) {
      if (event.kind === 'close' && event.localName === 'styleSheet') {
        closeStart = event.start
        const colon = event.name.indexOf(':')
        rootPrefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      }
    }
    if (closeStart === -1) {
      throw new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
        part: 'xl/styles.xml',
      })
    }

    const element = wanted.replace(/^<xf/, `<${rootPrefix}xf`)
    const table = `<${rootPrefix}cellXfs count="1">${element}</${rootPrefix}cellXfs>`
    return {
      xml: `${stylesXml.slice(0, closeStart)}${table}${stylesXml.slice(closeStart)}`,
      index: 0,
    }
  }

  const prefixed = wanted.replace(/^<xf/, `<${formats.prefix}xf`)
  const existing = formats.elements.indexOf(prefixed)
  if (existing !== -1) return { xml: stylesXml, index: existing }

  const index = formats.elements.length
  const openTag = formats.openTag.replace(/count="\d+"/, `count="${index + 1}"`)

  if (formats.selfClosing) {
    const opened = `${openTag.slice(0, -2)}>`
    return {
      xml:
        stylesXml.slice(0, formats.openStart) +
        `${opened}${prefixed}</${formats.prefix}cellXfs>` +
        stylesXml.slice(formats.openEnd),
      index,
    }
  }

  const head = stylesXml.slice(0, formats.openStart) + openTag
  const body = stylesXml.slice(formats.openEnd, formats.insertAt)
  return { xml: `${head}${body}${prefixed}${stylesXml.slice(formats.insertAt)}`, index }
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

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'numFmts' && event.selfClosing) {
      const opened = `${xml.slice(event.start, event.end - 2)}>`
      return `${xml.slice(0, event.start)}${opened}${element}</${prefix}numFmts>${xml.slice(event.end)}`
    }
    if (event.kind === 'close' && event.localName === 'numFmts') {
      return xml.slice(0, event.start) + element + xml.slice(event.start)
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
  for (let index = 0; index < parsed.cellFormats.length; index++) {
    if (numberFormatOf(parsed, index) === formatCode) return { xml: stylesXml, index }
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

  return applyFormatId(withFormat, basedOn, formatId)
}
