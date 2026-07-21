import { type Container, readContainer } from './container.js'
import { XlsxError } from './errors.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { readXml } from './xml.js'

const OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

const ROOT_RELATIONSHIPS = '_rels/.rels'

export type SheetState = 'visible' | 'hidden' | 'veryHidden'

export interface SheetRef {
  readonly name: string
  readonly sheetId: string
  /** Package path, already resolved. */
  readonly path: string
  readonly state: SheetState
}

export interface WorkbookPart {
  readonly sheets: readonly SheetRef[]
  /** The alternate epoch, where serials count from 1904 rather than 1900. */
  readonly date1904: boolean
  readonly container: Container
}

function partText(container: Container, path: string): string {
  const bytes = container.parts.get(path)
  if (bytes === undefined)
    throw new XlsxError('missing-part', `Missing part ${path}`, { part: path })
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch (cause) {
    throw new XlsxError('unreadable-part', `Part ${path} is not valid utf-8`, { part: path, cause })
  }
}

function findWorkbookPath(container: Container): string {
  if (!container.parts.has(ROOT_RELATIONSHIPS)) {
    throw new XlsxError('missing-part', `Missing part ${ROOT_RELATIONSHIPS}`, {
      part: ROOT_RELATIONSHIPS,
    })
  }

  for (const relationship of readRelationships(partText(container, ROOT_RELATIONSHIPS)).values()) {
    if (relationship.type === OFFICE_DOCUMENT) return resolveTarget('', relationship.target)
  }

  throw new XlsxError('missing-part', 'Package declares no workbook part', {
    part: ROOT_RELATIONSHIPS,
  })
}

function toSheetState(value: string | undefined): SheetState {
  if (value === 'hidden') return 'hidden'
  if (value === 'veryHidden') return 'veryHidden'
  return 'visible'
}

const isTrue = (value: string | undefined) => value === '1' || value === 'true'

/** The namespace prefix is the file's choice, so `r` cannot be assumed. */
function relationshipId(attributes: ReadonlyMap<string, string>): string | undefined {
  const direct = attributes.get('r:id')
  if (direct !== undefined) return direct

  for (const [name, value] of attributes) {
    if (name === 'id' || name.endsWith(':id')) return value
  }
  return undefined
}

export function readWorkbookPart(bytes: Uint8Array): WorkbookPart {
  const container = readContainer(bytes)
  const workbookPath = findWorkbookPath(container)
  const workbookXml = partText(container, workbookPath)

  const relationshipsPath = workbookPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const relationships = container.parts.has(relationshipsPath)
    ? readRelationships(partText(container, relationshipsPath))
    : new Map()

  const sheets: SheetRef[] = []
  let date1904 = false

  for (const event of readXml(workbookXml)) {
    if (event.kind !== 'open') continue

    if (event.localName === 'workbookPr') {
      date1904 = isTrue(event.attributes.get('date1904'))
      continue
    }
    if (event.localName !== 'sheet') continue

    const id = relationshipId(event.attributes)
    const relationship = id === undefined ? undefined : relationships.get(id)
    // Files exist that name a sheet with no usable relationship.
    if (relationship === undefined || relationship.external) continue

    sheets.push({
      name: event.attributes.get('name') ?? '',
      sheetId: event.attributes.get('sheetId') ?? '',
      path: resolveTarget(workbookPath, relationship.target),
      state: toSheetState(event.attributes.get('state')),
    })
  }

  return { sheets, date1904, container }
}
