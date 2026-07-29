import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  appendCommentsPart,
  appendVmlShapes,
  buildCommentsPart,
  buildVmlDrawing,
  withoutComment,
  withoutNoteShape,
  readComments,
} from './comments.js'
import { XlsxError } from './errors.js'

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
