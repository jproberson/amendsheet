import { isDateFormat, readStyles } from './styles.js'
import { readXml } from './xml.js'

/** Built in short date format, so no custom numFmt has to be written. */
const SHORT_DATE_FORMAT_ID = 14

const DEFAULT_DATE_XF = `<xf numFmtId="${SHORT_DATE_FORMAT_ID}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`

export interface DateStyle {
  readonly xml: string
  /** Index a cell should carry in its `s` attribute to display as a date. */
  readonly index: number
}

interface CellFormats {
  readonly elements: string[]
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
  const elements: string[] = []

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.name === 'cellXfs') {
      openTag = xml.slice(event.start, event.end)
      openStart = event.start
      openEnd = event.end
      selfClosing = event.selfClosing
      inside = !event.selfClosing
      if (selfClosing) insertAt = event.end
      continue
    }
    if (event.kind === 'close' && event.name === 'cellXfs') {
      insertAt = event.start
      inside = false
      continue
    }
    if (inside && event.kind === 'open' && event.name === 'xf') {
      elements.push(xml.slice(event.start, event.end))
    }
  }

  if (openStart === -1) return undefined
  return { elements, openTag, openStart, openEnd, insertAt, selfClosing }
}

function asDateFormat(element: string): string {
  const withFormat = element.includes('numFmtId="')
    ? element.replace(/numFmtId="\d+"/, `numFmtId="${SHORT_DATE_FORMAT_ID}"`)
    : element.replace(/^<xf/, `<xf numFmtId="${SHORT_DATE_FORMAT_ID}"`)

  return withFormat.includes('applyNumberFormat=')
    ? withFormat.replace(/applyNumberFormat="\d+"/, 'applyNumberFormat="1"')
    : withFormat.replace(/\/?>$/, ' applyNumberFormat="1"/>')
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

  const formats = readCellFormats(stylesXml)
  const wanted =
    basedOn === undefined
      ? DEFAULT_DATE_XF
      : asDateFormat(formats?.elements[basedOn] ?? DEFAULT_DATE_XF)

  if (formats === undefined) {
    const closing = stylesXml.lastIndexOf('</styleSheet>')
    const table = `<cellXfs count="1">${wanted}</cellXfs>`
    return {
      xml: `${stylesXml.slice(0, closing)}${table}${stylesXml.slice(closing)}`,
      index: 0,
    }
  }

  const existing = formats.elements.indexOf(wanted)
  if (existing !== -1) return { xml: stylesXml, index: existing }

  const index = formats.elements.length
  const openTag = formats.openTag.replace(/count="\d+"/, `count="${index + 1}"`)

  if (formats.selfClosing) {
    const opened = `${openTag.slice(0, -2)}>`
    return {
      xml:
        stylesXml.slice(0, formats.openStart) +
        `${opened}${wanted}</cellXfs>` +
        stylesXml.slice(formats.openEnd),
      index,
    }
  }

  const head = stylesXml.slice(0, formats.openStart) + openTag
  const body = stylesXml.slice(formats.openEnd, formats.insertAt)
  return { xml: `${head}${body}${wanted}${stylesXml.slice(formats.insertAt)}`, index }
}
