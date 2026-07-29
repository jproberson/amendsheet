import { readXml } from './xml.js'

/** The document's core properties, from `docProps/core.xml`. Every field is
 * optional; a workbook may set any subset. Dates are read and written in UTC. */
export interface DocumentProperties {
  readonly title?: string
  readonly subject?: string
  readonly creator?: string
  readonly keywords?: string
  readonly description?: string
  readonly lastModifiedBy?: string
  readonly created?: Date
  readonly modified?: Date
}

// localName in core.xml -> the field it maps to, and whether it holds a date.
const TEXT_FIELDS = [
  ['title', 'title'],
  ['subject', 'subject'],
  ['creator', 'creator'],
  ['keywords', 'keywords'],
  ['description', 'description'],
  ['lastModifiedBy', 'lastModifiedBy'],
] as const
const DATE_FIELDS = [
  ['created', 'created'],
  ['modified', 'modified'],
] as const

/** Reads the core properties, keeping only the fields this models. */
export function readCoreProperties(xml: string): DocumentProperties {
  const result: {
    -readonly [K in keyof DocumentProperties]?: DocumentProperties[K]
  } = {}
  let current: string | undefined
  let text = ''
  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      current = event.localName
      text = ''
    } else if (event.kind === 'text' && current !== undefined) {
      text += event.text
    } else if (event.kind === 'close' && current === event.localName) {
      for (const [name, field] of TEXT_FIELDS) {
        if (current === name && text !== '') result[field] = text
      }
      for (const [name, field] of DATE_FIELDS) {
        if (current === name) {
          const date = new Date(text)
          if (!Number.isNaN(date.getTime())) result[field] = date
        }
      }
      current = undefined
    }
  }
  return result
}

const escapeXml = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** W3CDTF in UTC, no fractional seconds — the form Excel writes. */
const w3cdtf = (date: Date): string => date.toISOString().replace(/\.\d{3}Z$/, 'Z')

const CORE_NAMESPACES =
  ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
  ' xmlns:dc="http://purl.org/dc/elements/1.1/"' +
  ' xmlns:dcterms="http://purl.org/dc/terms/"' +
  ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'

// The prefixed element each field is written as. dc: and cp: split the way Excel
// writes them; the dcterms dates carry the W3CDTF type marker.
const ELEMENT_FOR: Record<keyof DocumentProperties, { open: string; close: string }> = {
  title: { open: '<dc:title>', close: '</dc:title>' },
  subject: { open: '<dc:subject>', close: '</dc:subject>' },
  creator: { open: '<dc:creator>', close: '</dc:creator>' },
  keywords: { open: '<cp:keywords>', close: '</cp:keywords>' },
  description: { open: '<dc:description>', close: '</dc:description>' },
  lastModifiedBy: { open: '<cp:lastModifiedBy>', close: '</cp:lastModifiedBy>' },
  created: { open: '<dcterms:created xsi:type="dcterms:W3CDTF">', close: '</dcterms:created>' },
  modified: { open: '<dcterms:modified xsi:type="dcterms:W3CDTF">', close: '</dcterms:modified>' },
}

const FIELD_LOCAL_NAME: Record<keyof DocumentProperties, string> = {
  title: 'title',
  subject: 'subject',
  creator: 'creator',
  keywords: 'keywords',
  description: 'description',
  lastModifiedBy: 'lastModifiedBy',
  created: 'created',
  modified: 'modified',
}

function elementFor(field: keyof DocumentProperties, value: string | Date): string {
  const { open, close } = ELEMENT_FOR[field]
  const text = value instanceof Date ? w3cdtf(value) : escapeXml(value)
  return `${open}${text}${close}`
}

/** Replaces the element for `field`, or inserts it before the closing tag when
 * the part does not carry one, matching any namespace prefix on the existing
 * element so a field this does not touch is left exactly as it was. */
function withElement(xml: string, field: keyof DocumentProperties, value: string | Date): string {
  const element = elementFor(field, value)
  const local = FIELD_LOCAL_NAME[field]
  const existing = new RegExp(
    `<(?:\\w+:)?${local}\\b[^>]*>[\\s\\S]*?</(?:\\w+:)?${local}>|<(?:\\w+:)?${local}\\b[^>]*/>`,
  )
  if (existing.test(xml)) return xml.replace(existing, element)
  const close = xml.search(/<\/(?:\w+:)?coreProperties>/)
  return xml.slice(0, close) + element + xml.slice(close)
}

const EMPTY_CORE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<cp:coreProperties${CORE_NAMESPACES}></cp:coreProperties>`

/**
 * Writes the given properties into the core-properties part, opening a fresh one
 * when there is none. Only the fields provided are touched; the rest of an
 * existing part, including properties this does not model, is preserved.
 */
export function writeCoreProperties(
  existing: string | undefined,
  properties: DocumentProperties,
): string {
  let xml = existing ?? EMPTY_CORE
  for (const field of Object.keys(FIELD_LOCAL_NAME) as (keyof DocumentProperties)[]) {
    const value = properties[field]
    if (value !== undefined) xml = withElement(xml, field, value)
  }
  return xml
}
