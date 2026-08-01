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

const EMPTY_CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '</Types>'

test('applyContentTypes folds in the overrides and defaults declared on the draft', () => {
  const draft = createContainerDraft(containerOf({}), new Map())
  draft.declareOverride('xl/comments1.xml', 'application/comments')
  draft.declareDefault('png', 'image/png')
  const updated = draft.applyContentTypes(EMPTY_CONTENT_TYPES)
  assert.ok(
    updated.includes('<Override PartName="/xl/comments1.xml" ContentType="application/comments"/>'),
  )
  assert.ok(updated.includes('<Default Extension="png" ContentType="image/png"/>'))
})

test('applyContentTypes leaves a content-types part untouched when nothing was declared', () => {
  const draft = createContainerDraft(containerOf({}), new Map())
  assert.equal(draft.applyContentTypes(EMPTY_CONTENT_TYPES), EMPTY_CONTENT_TYPES)
})

test('declareOverride and declareDefault each register a part or extension once', () => {
  const draft = createContainerDraft(containerOf({}), new Map())
  draft.declareOverride('xl/tables/table1.xml', 'application/table')
  draft.declareOverride('xl/tables/table1.xml', 'application/table')
  draft.declareDefault('emf', 'image/emf')
  draft.declareDefault('emf', 'image/emf')
  const updated = draft.applyContentTypes(EMPTY_CONTENT_TYPES)
  assert.equal(updated.match(/PartName="\/xl\/tables\/table1.xml"/g)?.length, 1)
  assert.equal(updated.match(/Extension="emf"/g)?.length, 1)
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
