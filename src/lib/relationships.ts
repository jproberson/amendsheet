import { XlsxError } from './errors.js'
import { readXml } from './xml.js'

/** Parts refer to each other by relationship id rather than by path. */
export interface Relationship {
  readonly id: string
  readonly type: string
  readonly target: string
  /** External targets are URLs and name nothing inside the package. */
  readonly external: boolean
}

export function readRelationships(xml: string): ReadonlyMap<string, Relationship> {
  const relationships = new Map<string, Relationship>()

  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Relationship') continue

    const id = event.attributes.get('Id')
    if (id === undefined) throw new XlsxError('malformed-xml', 'Relationship is missing Id')

    const target = event.attributes.get('Target')
    if (target === undefined)
      throw new XlsxError('malformed-xml', `Relationship ${id} is missing Target`)

    relationships.set(id, {
      id,
      type: event.attributes.get('Type') ?? '',
      target,
      external: event.attributes.get('TargetMode') === 'External',
    })
  }

  return relationships
}

/** Targets are relative to the folder of the declaring part, unless absolute. */
export function resolveTarget(ownerPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)

  const ownerFolder = ownerPath.split('/').slice(0, -1)
  const segments = ownerPath === '' ? [] : ownerFolder

  for (const segment of target.split('/')) {
    if (segment === '.' || segment === '') continue
    if (segment === '..') segments.pop()
    else segments.push(segment)
  }

  return segments.join('/')
}
