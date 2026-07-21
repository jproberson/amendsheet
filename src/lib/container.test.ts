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
