import assert from 'node:assert/strict'
import { test } from 'node:test'
import type { Container } from './container.js'
import { createContainerDraft, withRelationship } from './container-draft.js'

const encode = (text: string) => new TextEncoder().encode(text)

const containerOf = (parts: Record<string, string>): Container => ({
  parts: new Map(Object.entries(parts).map(([path, text]) => [path, encode(text)])),
})

test('text prefers a staged change over the original', () => {
  const changes = new Map<string, Uint8Array | null>()
  const draft = createContainerDraft(containerOf({ 'a.xml': '<a/>' }), changes)
  assert.equal(draft.text('a.xml'), '<a/>')
  draft.setBytes('a.xml', encode('<b/>'))
  assert.equal(draft.text('a.xml'), '<b/>')
  assert.equal(draft.original('a.xml'), '<a/>')
})

test('text is undefined for a part that is nowhere', () => {
  const draft = createContainerDraft(containerOf({}), new Map())
  assert.equal(draft.text('missing.xml'), undefined)
  assert.equal(draft.original('missing.xml'), undefined)
})

test('has covers both the container and staged changes', () => {
  const changes = new Map<string, Uint8Array | null>()
  const draft = createContainerDraft(containerOf({ 'x.xml': '<x/>' }), changes)
  draft.setBytes('y.xml', encode('<y/>'))
  assert.ok(draft.has('x.xml'))
  assert.ok(draft.has('y.xml'))
  assert.equal(draft.has('z.xml'), false)
})

test('freeNumber skips numbers taken by the container or a staged change', () => {
  const changes = new Map<string, Uint8Array | null>()
  const draft = createContainerDraft(containerOf({ 'xl/comments1.xml': '' }), changes)
  draft.setBytes('xl/comments2.xml', encode(''))
  assert.equal(
    draft.freeNumber((n) => `xl/comments${n}.xml`),
    3,
  )
})

test('relationshipTarget resolves an owner rel of a type', () => {
  const rels =
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://example/comments" Target="../comments1.xml"/>' +
    '</Relationships>'
  const draft = createContainerDraft(
    containerOf({ 'xl/worksheets/_rels/sheet1.xml.rels': rels }),
    new Map(),
  )
  assert.equal(
    draft.relationshipTarget('xl/worksheets/sheet1.xml', 'http://example/comments'),
    'xl/comments1.xml',
  )
  assert.equal(
    draft.relationshipTarget('xl/worksheets/sheet1.xml', 'http://example/other'),
    undefined,
  )
})

test('withRelationship appends a relationship and returns its fresh id', () => {
  const empty = undefined
  const first = withRelationship(empty, 'http://example/t', 'target.xml')
  assert.equal(first.id, 'rId1')
  const second = withRelationship(first.xml, 'http://example/u', 'other.xml')
  assert.equal(second.id, 'rId2')
  assert.ok(second.xml.includes('Target="target.xml"'))
  assert.ok(second.xml.includes('Target="other.xml"'))
})
