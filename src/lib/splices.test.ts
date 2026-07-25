import assert from 'node:assert/strict'
import { test } from 'node:test'

import { applySplices } from './splices.js'

test('applySplices replaces, inserts and cuts, in start order', () => {
  const text = 'abcdefg'
  assert.equal(
    applySplices(text, [
      { start: 5, end: 5, text: 'X' }, // insert, given out of order
      { start: 1, end: 3, text: 'Z' }, // replace bc
      { start: 4, end: 5, text: '' }, // cut e
    ]),
    'aZdXfg',
  )
})

test('applySplices with nothing to do returns the text', () => {
  assert.equal(applySplices('abc', []), 'abc')
})
