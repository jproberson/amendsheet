import assert from 'node:assert/strict'
import { test } from 'node:test'
import { assertWellFormed } from '../testing/invariants.js'
import { buildRichInline, flattenRuns, richTextOf, runFont } from './rich-text.js'

test('runFont reads a run font off its rPr', () => {
  assert.deepEqual(runFont('<rPr><b/><rFont val="Arial"/></rPr>'), { bold: true, name: 'Arial' })
})

test('runFont on an empty rPr is no font, not an empty object', () => {
  assert.equal(runFont('<rPr></rPr>'), undefined)
})

test('richTextOf treats a single unformatted run as plain text', () => {
  assert.equal(richTextOf([{ text: 'lonely' }]), undefined)
  assert.equal(richTextOf([]), undefined)
})

test('richTextOf keeps a single run that carries a font', () => {
  assert.deepEqual(richTextOf([{ text: 'x', font: { italic: true } }]), {
    runs: [{ text: 'x', font: { italic: true } }],
  })
})

test('flattenRuns joins every run text', () => {
  assert.equal(flattenRuns([{ text: 'a' }, { text: 'b', font: { bold: true } }]), 'ab')
})

test('buildRichInline emits a run per stretch, rPr only where a font is set', () => {
  const is = buildRichInline([{ text: 'a', font: { bold: true } }, { text: 'b' }], '')

  assert.equal(is, '<is><r><rPr><b/></rPr><t>a</t></r><r><t>b</t></r></is>')
  assertWellFormed(`<c>${is}</c>`, 'inline rich string')
})

test('buildRichInline preserves edge whitespace with xml:space', () => {
  assert.match(buildRichInline([{ text: ' pad ' }], ''), /<t xml:space="preserve"> pad <\/t>/)
})

test('buildRichInline prefixes every element to match the sheet', () => {
  const is = buildRichInline([{ text: 'x', font: { bold: true } }], 'x:')

  assert.equal(is, '<x:is><x:r><x:rPr><x:b/></x:rPr><x:t>x</x:t></x:r></x:is>')
})
