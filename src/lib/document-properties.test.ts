import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readCoreProperties, writeCoreProperties } from './document-properties.js'

const CORE =
  '<?xml version="1.0"?>' +
  '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc" xmlns:dcterms="dcterms" xmlns:xsi="xsi">' +
  '<dc:title>Report</dc:title><dc:creator>Ada</dc:creator>' +
  '<cp:revision>4</cp:revision>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>' +
  '</cp:coreProperties>'

test('readCoreProperties reads text and date fields', () => {
  const props = readCoreProperties(CORE)
  assert.equal(props.title, 'Report')
  assert.equal(props.creator, 'Ada')
  assert.deepEqual(props.created, new Date('2024-01-02T03:04:05Z'))
  assert.equal(props.subject, undefined)
})

test('writeCoreProperties builds a fresh part with only the given fields', () => {
  const xml = writeCoreProperties(undefined, { title: 'A & B', creator: 'Grace' })
  assert.match(xml, /<dc:title>A &amp; B<\/dc:title>/)
  assert.match(xml, /<dc:creator>Grace<\/dc:creator>/)
  assert.doesNotMatch(xml, /subject/)
  assert.deepEqual(readCoreProperties(xml), { title: 'A & B', creator: 'Grace' })
})

test('writeCoreProperties replaces a field in place and keeps the rest', () => {
  const xml = writeCoreProperties(CORE, {
    title: 'New',
    modified: new Date('2025-06-07T08:09:10Z'),
  })
  assert.match(xml, /<dc:title>New<\/dc:title>/)
  assert.doesNotMatch(xml, /<dc:title>Report<\/dc:title>/) // replaced, not doubled
  assert.equal((xml.match(/<dc:title>/g) ?? []).length, 1)
  assert.match(xml, /<dc:creator>Ada<\/dc:creator>/) // untouched field kept
  assert.match(xml, /<cp:revision>4<\/cp:revision>/) // unmodelled field preserved
  assert.match(
    xml,
    /<dcterms:modified xsi:type="dcterms:W3CDTF">2025-06-07T08:09:10Z<\/dcterms:modified>/,
  )
})
