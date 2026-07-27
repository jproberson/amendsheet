import assert from 'node:assert/strict'
import { test } from 'node:test'
import { buildCommentsPart, readComments } from './comments.js'

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

test('buildCommentsPart writes each note under one empty author, escaping the text', () => {
  const xml = buildCommentsPart(new Map([['A1', 'a < b & c']]))
  assert.match(xml, /<authors><author\/><\/authors>/)
  assert.match(
    xml,
    /<comment ref="A1" authorId="0"><text><t xml:space="preserve">a &lt; b &amp; c<\/t><\/text><\/comment>/,
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
