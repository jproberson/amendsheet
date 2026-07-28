import { escapeSheetName } from './add-sheet.js'
import { canonicalReference, parseReference } from './reference.js'
import { readRelationships } from './relationships.js'
import { type Splice, applySplices } from './splices.js'
import { readXml } from './xml.js'

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HYPERLINK_TYPE = `${RELATIONSHIPS_NS}/hyperlink`
const PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships'

// The elements CT_Worksheet places after hyperlinks, so a fresh hyperlinks lands
// before the first of them the sheet has, and before </worksheet> when it has none.
const AFTER_HYPERLINKS = new Set([
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
])

export interface HyperlinkEntry {
  readonly reference: string
  /** Set for a link out of the package; the relationship carries the URL. */
  readonly relationshipId?: string
  /** Set for a link within the workbook — a cell or a defined name. */
  readonly location?: string
  readonly tooltip?: string
}

/** Writes the hyperlinks into a sheet, merging with any it has and replacing one on the same cell. */
export function withHyperlinks(sheetXml: string, entries: readonly HyperlinkEntry[]): string {
  if (entries.length === 0) return sheetXml
  const replaced = new Set(entries.map((entry) => entry.reference))
  const external = entries.some((entry) => entry.relationshipId !== undefined)

  let prefix = ''
  let worksheet: { start: number; end: number } | undefined
  let container: { start: number; end: number; selfClosing: boolean } | undefined
  const stale: { start: number; end: number }[] = []
  let afterStart = -1
  let worksheetCloseStart = -1
  let hyperlinksCloseStart = -1
  let insideContainer = false

  for (const event of readXml(sheetXml)) {
    if (event.kind === 'close') {
      if (event.localName === 'hyperlinks') {
        hyperlinksCloseStart = event.start
        insideContainer = false
      } else if (event.localName === 'worksheet') worksheetCloseStart = event.start
      continue
    }
    if (event.kind !== 'open') continue
    if (event.localName === 'worksheet') {
      const colon = event.name.indexOf(':')
      prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      worksheet = { start: event.start, end: event.end }
    } else if (event.localName === 'hyperlinks') {
      container = { start: event.start, end: event.end, selfClosing: event.selfClosing }
      insideContainer = !event.selfClosing
    } else if (event.localName === 'hyperlink' && insideContainer) {
      const reference = event.attributes.get('ref')
      if (reference !== undefined && replaced.has(reference))
        stale.push({ start: event.start, end: event.end })
    } else if (afterStart === -1 && !insideContainer && AFTER_HYPERLINKS.has(event.localName)) {
      afterStart = event.start
    }
  }

  const child = (entry: HyperlinkEntry): string => {
    let attributes = `ref="${entry.reference}"`
    if (entry.relationshipId !== undefined) attributes += ` r:id="${entry.relationshipId}"`
    if (entry.location !== undefined) attributes += ` location="${escapeSheetName(entry.location)}"`
    if (entry.tooltip !== undefined) attributes += ` tooltip="${escapeSheetName(entry.tooltip)}"`
    return `<${prefix}hyperlink ${attributes}/>`
  }
  const children = entries.map(child).join('')

  const splices: Splice[] = stale.map((span) => ({ start: span.start, end: span.end, text: '' }))

  // An external link is addressed through r:id, so the sheet must declare that
  // namespace; a blank sheet this library created does not.
  if (external && worksheet !== undefined) {
    const tag = sheetXml.slice(worksheet.start, worksheet.end)
    if (!tag.includes('xmlns:r=')) {
      splices.push({
        start: worksheet.start,
        end: worksheet.end,
        text: tag.replace(/^<([^\s/>]+)/, `<$1 xmlns:r="${RELATIONSHIPS_NS}"`),
      })
    }
  }

  const wrapped = `<${prefix}hyperlinks>${children}</${prefix}hyperlinks>`
  if (container === undefined) {
    const anchor =
      afterStart !== -1
        ? afterStart
        : worksheetCloseStart !== -1
          ? worksheetCloseStart
          : sheetXml.length
    splices.push({ start: anchor, end: anchor, text: wrapped })
  } else if (container.selfClosing) {
    splices.push({ start: container.start, end: container.end, text: wrapped })
  } else {
    splices.push({ start: hyperlinksCloseStart, end: hyperlinksCloseStart, text: children })
  }

  return applySplices(sheetXml, splices)
}

/** Adds each external link's relationship to a sheet's rels part, creating it when absent. */
export function withHyperlinkRelationships(
  relsXml: string | undefined,
  links: readonly { id: string; url: string }[],
): string {
  const elements = links
    .map(
      (link) =>
        `<Relationship Id="${link.id}" Type="${HYPERLINK_TYPE}" Target="${escapeSheetName(link.url)}" TargetMode="External"/>`,
    )
    .join('')
  if (relsXml === undefined) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_RELATIONSHIPS}">${elements}</Relationships>`
  }
  return relsXml.replace(/<\/(\w+:)?Relationships>/, `${elements}$&`)
}

/**
 * Where a cell links to: a URL out of the package, or a `location` within the
 * workbook — a cell reference like `Sheet2!A1` or a defined name. `tooltip` is the
 * hover text.
 */
export type Hyperlink =
  | { readonly url: string; readonly tooltip?: string }
  | { readonly location: string; readonly tooltip?: string }

export interface SheetHyperlinkWrite {
  readonly sheetXml: string
  /** Present only when a link needs an external relationship in the sheet's rels. */
  readonly relsXml?: string
}

/**
 * Writes a sheet's pending hyperlinks into its XML. An external URL takes a fresh
 * relationship id past the highest the rels part already uses; an internal
 * location is written inline and needs none.
 */
export function writeSheetHyperlinks(
  sheetXml: string,
  existingRels: string | undefined,
  links: ReadonlyMap<string, Hyperlink>,
): SheetHyperlinkWrite {
  let nextId = 0
  for (const match of (existingRels ?? '').matchAll(/Id="rId(\d+)"/g)) {
    nextId = Math.max(nextId, Number(match[1]))
  }
  const entries: HyperlinkEntry[] = []
  const externalRels: { id: string; url: string }[] = []
  for (const [reference, target] of links) {
    if ('url' in target) {
      const id = `rId${++nextId}`
      entries.push({ reference, relationshipId: id, tooltip: target.tooltip })
      externalRels.push({ id, url: target.url })
    } else {
      entries.push({ reference, location: target.location, tooltip: target.tooltip })
    }
  }
  const written = withHyperlinks(sheetXml, entries)
  if (externalRels.length === 0) return { sheetXml: written }
  return { sheetXml: written, relsXml: withHyperlinkRelationships(existingRels, externalRels) }
}

/**
 * Reads a sheet's hyperlinks, keyed by the cell each anchors at — the top-left of
 * the ref, so a range link lands on the cell a reader clicks first. An `r:id`
 * link resolves to its external URL through the sheet's rels; a `location` link is
 * inline. A link whose relationship is missing is dropped rather than reported
 * pointing nowhere.
 */
export function readSheetHyperlinks(
  sheetXml: string,
  relsXml: string | undefined,
  part: string,
): ReadonlyMap<string, Hyperlink> {
  const links = new Map<string, Hyperlink>()
  const relationships = relsXml === undefined ? undefined : readRelationships(relsXml, part)
  let inside = false
  for (const event of readXml(sheetXml)) {
    if (event.kind === 'close') {
      if (event.localName === 'hyperlinks') inside = false
      continue
    }
    if (event.kind !== 'open') continue
    if (event.localName === 'hyperlinks') {
      inside = !event.selfClosing
      continue
    }
    if (!inside || event.localName !== 'hyperlink') continue
    const ref = event.attributes.get('ref')
    if (ref === undefined) continue
    const colon = ref.indexOf(':')
    const anchor = colon === -1 ? ref : ref.slice(0, colon)
    const key = canonicalReference(parseReference(anchor)) ?? anchor
    const tooltip = event.attributes.get('tooltip')
    const relationshipId = event.attributes.get('r:id')
    if (relationshipId !== undefined) {
      const url = relationships?.get(relationshipId)?.target
      if (url !== undefined) links.set(key, tooltip === undefined ? { url } : { url, tooltip })
    } else {
      const location = event.attributes.get('location')
      if (location !== undefined) {
        links.set(key, tooltip === undefined ? { location } : { location, tooltip })
      }
    }
  }
  return links
}
