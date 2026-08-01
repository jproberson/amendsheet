import assert from 'node:assert/strict'
import { test } from 'node:test'
import { contributeComments } from './comments.js'
import type { Container } from './container.js'
import { createContainerDraft } from './container-draft.js'
import { contributeHyperlinks } from './hyperlinks.js'
import { contributeImages } from './images.js'
import { contributeTables } from './tables.js'

const encode = (text: string) => new TextEncoder().encode(text)
const PATH = 'xl/worksheets/sheet1.xml'
const RELS = 'xl/worksheets/_rels/sheet1.xml.rels'
const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>'

const freshSheet = () => {
  const changes = new Map<string, Uint8Array | null>()
  const container: Container = { parts: new Map([[PATH, encode('<worksheet></worksheet>')]]) }
  return { changes, draft: createContainerDraft(container, changes) }
}

test('contributeComments opens a comments part and a VML drawing wired to the sheet', () => {
  const { draft } = freshSheet()
  const result = contributeComments(
    draft,
    new Map([[PATH, new Map([['A1', 'hello']])]]),
    new Map(),
    new Set(),
  )
  assert.deepEqual(result.commentParts, ['xl/comments1.xml'])
  assert.deepEqual(result.vmlDrawingParts, ['xl/drawings/vmlDrawing1.vml'])
  assert.ok(draft.text('xl/comments1.xml')?.includes('hello'))
  assert.ok(draft.text(PATH)?.includes('legacyDrawing'))
  assert.ok(
    draft.relationshipTarget(
      PATH,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    ),
  )
})

test('contributeImages writes a media part and a drawing, returning the extension', () => {
  const { draft } = freshSheet()
  const result = contributeImages(
    draft,
    new Map([
      [
        PATH,
        [
          {
            bytes: encode('img'),
            type: 'png',
            from: { column: 0, row: 0 },
            to: { column: 1, row: 1 },
          },
        ],
      ],
    ]),
    new Set(),
  )
  assert.deepEqual(result.drawingParts, ['xl/drawings/drawing1.xml'])
  assert.deepEqual([...result.imageExtensions], ['png'])
  assert.ok(draft.has('xl/media/image1.png'))
  assert.ok(draft.text('xl/drawings/drawing1.xml')?.includes('twoCellAnchor'))
  assert.ok(draft.text(PATH)?.includes('<drawing '))
})

test('contributeTables writes a table part, lists it on the sheet and declares its type', () => {
  const { draft } = freshSheet()
  contributeTables(
    draft,
    new Map([
      [PATH, [{ name: 'T1', ref: 'A1:B2', columns: ['a', 'b'], style: 'TableStyleMedium2' }]],
    ]),
    new Set(),
  )
  assert.ok(draft.text('xl/tables/table1.xml')?.includes('displayName="T1"'))
  assert.ok(draft.text(PATH)?.includes('tableParts'))
  assert.ok(draft.applyContentTypes(CONTENT_TYPES).includes('PartName="/xl/tables/table1.xml"'))
})

test('contributeHyperlinks writes the link inline and its external relationship', () => {
  const { draft } = freshSheet()
  contributeHyperlinks(
    draft,
    new Map([[PATH, new Map([['A1', { url: 'https://example.com' }]])]]),
    new Map(),
    new Set(),
  )
  assert.ok(draft.text(PATH)?.includes('hyperlink'))
  assert.ok(draft.text(RELS)?.includes('example.com'))
})

test('a removed sheet contributes nothing', () => {
  const { draft, changes } = freshSheet()
  const result = contributeComments(
    draft,
    new Map([[PATH, new Map([['A1', 'x']])]]),
    new Map(),
    new Set([PATH]),
  )
  assert.deepEqual(result.commentParts, [])
  assert.equal(changes.size, 0)
})
