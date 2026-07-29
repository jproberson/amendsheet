import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  appendCommentsPart,
  appendVmlShapes,
  buildCommentsPart,
  buildVmlDrawing,
  readComments,
  shiftComments,
  shiftNoteShapes,
  withoutComment,
  withoutNoteShape,
} from './comments.js'
import { XlsxError } from './errors.js'
import type { ShiftSpec } from './shift.js'

const rowSpec = (at: number, delta: number): ShiftSpec => ({
  axis: 'row',
  at,
  delta,
  editedSheet: 'Sheet1',
  onCurrentSheet: true,
})

test('readComments maps each cell to the joined text of its runs', () => {
  const xml =
    '<comments><authors><author>Ada</author></authors><commentList>' +
    '<comment ref="B2" authorId="0"><text><r><t>Check </t></r><r><t>this</t></r></text></comment>' +
    '<comment ref="C3" authorId="0"><text><t>Plain &amp; short</t></text></comment>' +
    '</commentList></comments>'
  const comments = readComments(xml)
  assert.equal(comments.get('B2'), 'Check this')
  assert.equal(comments.get('C3'), 'Plain & short')
  assert.equal(comments.size, 2)
})

test('buildCommentsPart wraps each note text in a run, escaping it', () => {
  const xml = buildCommentsPart(new Map([['A1', 'a < b & c']]))
  assert.match(xml, /<authors><author\/><\/authors>/)
  assert.match(
    xml,
    /<comment ref="A1" authorId="0"><text><r><t xml:space="preserve">a &lt; b &amp; c<\/t><\/r><\/text><\/comment>/,
  )
})

test('buildVmlDrawing gives each note one hidden shape anchored to its cell', () => {
  const vml = buildVmlDrawing(['A1', 'C3'])
  assert.equal(vml.match(/<v:shapetype /g)?.length, 1)
  assert.match(vml, /<v:shape id="_x0000_s1025"/)
  assert.match(vml, /<v:shape id="_x0000_s1026"/)
  assert.equal(vml.match(/ObjectType="Note"/g)?.length, 2)
  assert.equal(vml.match(/visibility:hidden/g)?.length, 2)
  assert.match(vml, /<x:Row>0<\/x:Row><x:Column>0<\/x:Column>/)
  assert.match(vml, /<x:Row>2<\/x:Row><x:Column>2<\/x:Column>/)
})

test('buildVmlDrawing with no notes still closes to a well-formed part', () => {
  const vml = buildVmlDrawing([])
  assert.match(vml, /<v:shapetype /)
  assert.doesNotMatch(vml, /<v:shape /)
  assert.match(vml, /<\/xml>$/)
})

test('appendCommentsPart adds notes without disturbing the rich text already there', () => {
  const existing =
    '<?xml version="1.0"?>\n<comments><authors><author>Ada</author></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>bold note</t></r></text></comment>' +
    '</commentList></comments>'
  const merged = appendCommentsPart(existing, new Map([['B2', 'plain']]))
  // The existing rich comment survives byte for byte.
  assert.match(merged, /<r><rPr><b\/><\/rPr><t>bold note<\/t><\/r>/)
  assert.match(merged, /<comment ref="B2" authorId="0"><text><r><t xml:space="preserve">plain</)
  const read = readComments(merged)
  assert.equal(read.get('A1'), 'bold note')
  assert.equal(read.get('B2'), 'plain')
})

test('appendCommentsPart on a malformed part is rejected', () => {
  assert.throws(
    () => appendCommentsPart('<comments/>', new Map([['A1', 'x']])),
    (error: unknown) => error instanceof XlsxError && error.code === 'invalid-content',
  )
})

test('appendVmlShapes gives new shapes ids and z-indexes past the existing ones', () => {
  const existing = buildVmlDrawing(['A1']) // one shape at _x0000_s1025, z-index 1
  const merged = appendVmlShapes(existing, ['C3'])
  assert.match(merged, /<v:shape id="_x0000_s1025"/) // kept
  assert.match(merged, /<v:shape id="_x0000_s1026"[^>]*z-index:2/) // appended past it
  assert.match(merged, /<x:Row>2<\/x:Row><x:Column>2<\/x:Column>/)
  assert.match(merged, /<\/xml>$/)
})

test('appendVmlShapes on a malformed drawing is rejected', () => {
  assert.throws(
    () => appendVmlShapes('<not-a-drawing/>', ['A1']),
    (error: unknown) => error instanceof XlsxError && error.code === 'invalid-content',
  )
})

test('a built comments part reads back to the same text', () => {
  const built = buildCommentsPart(
    new Map([
      ['A1', 'hello'],
      ['Z9', 'world'],
    ]),
  )
  const read = readComments(built)
  assert.equal(read.get('A1'), 'hello')
  assert.equal(read.get('Z9'), 'world')
})

test('withoutComment removes the comment for a cell and keeps the rest', () => {
  const xml = buildCommentsPart(
    new Map([
      ['A1', 'first'],
      ['C3', 'second'],
    ]),
  )
  const out = withoutComment(xml, 'A1')
  assert.doesNotMatch(out, /ref="A1"/)
  assert.match(out, /<comment ref="C3"/)
  assert.equal(withoutComment(xml, 'Z9'), xml) // absent ref, unchanged
})

test('withoutNoteShape removes the shape anchored at a cell, keeping others', () => {
  const vml = buildVmlDrawing(['A1', 'C3']) // shapes at (row0,col0) 0,0 and 2,2
  const out = withoutNoteShape(vml, 0, 0) // A1
  assert.doesNotMatch(out, /<x:Row>0<\/x:Row>/)
  assert.match(out, /<x:Row>2<\/x:Row><x:Column>2<\/x:Column>/) // C3 stays
  assert.equal(out.match(/<v:shape /g)?.length, 1)
  assert.equal(withoutNoteShape(vml, 9, 9), vml) // no shape there, unchanged
})

test('shiftComments moves, drops and leaves refs by what the edit does to their cell', () => {
  const xml =
    '<comments><authors><author/></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><r><t>above</t></r></text></comment>' +
    '<comment ref="A3" authorId="0"><text><r><t>below</t></r></text></comment>' +
    '</commentList></comments>'
  // Insert one row at row 2: A1 is above and unmoved, A3 slides to A4.
  const inserted = shiftComments(xml, rowSpec(2, 1))
  assert.match(inserted, /ref="A1"/)
  assert.match(inserted, /ref="A4"/)
  // Delete row 3: A3's cell is gone, so its whole comment goes; A1 stays.
  const deleted = shiftComments(xml, rowSpec(3, -1))
  assert.match(deleted, /ref="A1"/)
  assert.doesNotMatch(deleted, /ref="A[34]"/)
  assert.doesNotMatch(deleted, /below/)
})

test('shiftComments leaves a ref it cannot parse alone', () => {
  const xml =
    '<comments><commentList>' +
    '<comment ref="not-a-ref" authorId="0"><text><r><t>x</t></r></text></comment>' +
    '</commentList></comments>'
  assert.equal(shiftComments(xml, rowSpec(1, 1)), xml)
})

test('shiftNoteShapes moves a box row anchor and corners, dropping a deleted one', () => {
  const vml = buildVmlDrawing(['A2', 'A5']) // rows 1 and 4 zero-based
  // Insert a row at row 2: A2's box moves down one, A5's moves down one too.
  const inserted = shiftNoteShapes(vml, rowSpec(2, 1))
  assert.match(inserted, /<x:Row>2<\/x:Row>/) // A2 -> row index 2
  assert.match(inserted, /<x:Row>5<\/x:Row>/) // A5 -> row index 5
  const movedAnchor = inserted.match(/<x:Anchor>([^<]*)<\/x:Anchor>/)?.[1]?.split(',') ?? []
  assert.equal(movedAnchor[2]?.trim(), '2') // top-row corner followed the cell
  // Delete row 2: A2's cell is removed, so its shape goes; A5 slides up.
  const deleted = shiftNoteShapes(vml, rowSpec(2, -1))
  assert.equal(deleted.match(/<v:shape /g)?.length, 1)
  assert.match(deleted, /<x:Row>3<\/x:Row>/) // A5 -> row index 3
})

test('shiftNoteShapes leaves a shape with no cell and a malformed anchor be', () => {
  const noCell = '<xml><v:shape id="_x0000_s1"><v:textbox/></v:shape></xml>'
  assert.equal(shiftNoteShapes(noCell, rowSpec(1, 1)), noCell)
  const badAnchor =
    '<xml><v:shape><x:ClientData ObjectType="Note">' +
    '<x:Anchor>1, 2, 3</x:Anchor><x:Row>4</x:Row><x:Column>0</x:Column>' +
    '</x:ClientData></v:shape></xml>'
  const out = shiftNoteShapes(badAnchor, rowSpec(1, 1))
  assert.match(out, /<x:Row>5<\/x:Row>/) // the row still moves
  assert.match(out, /<x:Anchor>1, 2, 3<\/x:Anchor>/) // the three-field anchor is left
})
