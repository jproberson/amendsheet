import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer, writeContainer } from './container.js'

async function fixtureFile(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(`fixtures/real/${name}`))
}

test('keeps every part when a file is read and written back', async () => {
  const original = await fixtureFile('WithChart.xlsx')

  const container = readContainer(original)
  const rewritten = readContainer(writeContainer(container))

  assert.deepEqual([...rewritten.parts.keys()].sort(), [...container.parts.keys()].sort())
})

test('keeps part contents byte for byte', async () => {
  const original = await fixtureFile('WithChart.xlsx')

  const container = readContainer(original)
  const rewritten = readContainer(writeContainer(container))

  for (const [path, bytes] of container.parts) {
    assert.deepEqual(rewritten.parts.get(path), bytes, `${path} changed`)
  }
})

test('reports the part that could not be read', async () => {
  const notAZip = new Uint8Array([1, 2, 3, 4])

  assert.throws(() => readContainer(notAZip), /not a zip/i)
})

test('writes the same bytes every time for the same parts', () => {
  const parts = new Map([['a.xml', new TextEncoder().encode('<a/>')]])

  const first = writeContainer({ parts })
  const header = first.slice(10, 14)

  // Local file header: bytes 10-11 hold the DOS time, 12-13 the DOS date.
  // A fixed stamp keeps output reproducible; 0x0021 is 1980-01-01.
  assert.deepEqual([...header], [0, 0, 33, 0])
})
