import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readPageMargins, readPageSetup, withPageMargins, withPageSetup } from './page.js'

const encode = (text: string) => new TextEncoder().encode(text)

test('withPageSetup inserts a fresh element before its successors', () => {
  const out = withPageSetup('<worksheet><sheetData/><tableParts count="1"/></worksheet>', {
    orientation: 'landscape',
    scale: 80,
  })
  assert.match(out, /<pageSetup orientation="landscape" scale="80"\/><tableParts/)
})

test('withPageSetup merges onto an existing element, keeping other attributes', () => {
  const out = withPageSetup(
    '<worksheet><sheetData/><pageSetup paperSize="9" orientation="portrait"/></worksheet>',
    { orientation: 'landscape' },
  )
  assert.match(out, /<pageSetup paperSize="9" orientation="landscape"\/>/)
})

test('withPageMargins writes all six, merging onto the ones present', () => {
  const fresh = withPageMargins('<worksheet><sheetData/></worksheet>', { left: 1 })
  assert.match(
    fresh,
    /<pageMargins left="1" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"\/>/,
  )
  const merged = withPageMargins(
    '<worksheet><sheetData/><pageMargins left="0.5" right="0.5" top="1" bottom="1" header="0.4" footer="0.4"/></worksheet>',
    { top: 2 },
  )
  assert.match(merged, /left="0.5" right="0.5" top="2" bottom="1" header="0.4" footer="0.4"/)
})

test('readPageSetup and readPageMargins read back what was written', () => {
  const setup = withPageSetup('<worksheet><sheetData/></worksheet>', {
    orientation: 'landscape',
    scale: 120,
  })
  assert.deepEqual(readPageSetup(encode(setup)), { orientation: 'landscape', scale: 120 })
  assert.deepEqual(readPageSetup(encode('<worksheet><sheetData/></worksheet>')), {})

  const margins = withPageMargins('<worksheet><sheetData/></worksheet>', { left: 1, footer: 0.2 })
  assert.deepEqual(readPageMargins(encode(margins)), {
    left: 1,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.2,
  })
})

test('withPageSetup inserts a missing attribute when merging, and handles a rootless fragment', () => {
  const added = withPageSetup(
    '<worksheet><sheetData/><pageSetup orientation="portrait"/></worksheet>',
    { scale: 50 }, // no scale on the existing element
  )
  assert.match(added, /<pageSetup scale="50" orientation="portrait"\/>/) // inserted after the tag name

  const rootless = withPageSetup('<sheetData/>', {}) // no orientation, no scale, no worksheet
  assert.equal(rootless, '<sheetData/><pageSetup/>')

  const orientationOnly = withPageSetup('<worksheet><sheetData/></worksheet>', {
    orientation: 'portrait',
  })
  assert.match(orientationOnly, /<pageSetup orientation="portrait"\/>/)
  const scaleOnly = withPageSetup('<worksheet><sheetData/></worksheet>', { scale: 75 })
  assert.match(scaleOnly, /<pageSetup scale="75"\/>/)
})

test('withPageMargins appends before a successor when the sheet has none', () => {
  const out = withPageMargins('<worksheet><sheetData/><pageSetup/></worksheet>', { header: 0.5 })
  assert.match(out, /<pageMargins [^>]*header="0.5"[^>]*\/><pageSetup\/>/)
})

test('readPageSetup ignores a non-standard orientation and a missing scale', () => {
  assert.deepEqual(
    readPageSetup(encode('<worksheet><sheetData/><pageSetup orientation="default"/></worksheet>')),
    {},
  )
})

test('readPageMargins reads only the sides present', () => {
  assert.deepEqual(
    readPageMargins(encode('<worksheet><sheetData/><pageMargins left="1" top="2"/></worksheet>')),
    { left: 1, top: 2 },
  )
})

test('page reads ignore a non-numeric value and setAttribute handles single quotes', () => {
  assert.deepEqual(
    readPageMargins(encode('<worksheet><sheetData/><pageMargins left="x" top="2"/></worksheet>')),
    { top: 2 }, // the unparseable left is dropped
  )
  assert.deepEqual(
    readPageSetup(encode('<worksheet><sheetData/><pageSetup scale="wat"/></worksheet>')),
    {},
  )
  const merged = withPageSetup(
    "<worksheet><sheetData/><pageSetup orientation='portrait'/></worksheet>",
    { orientation: 'landscape' },
  )
  assert.match(merged, /orientation="landscape"/)
})

test('withPageMargins fills defaults for sides an existing element omits', () => {
  const out = withPageMargins(
    '<worksheet><sheetData/><pageMargins left="0.5"/></worksheet>', // only left present
    { right: 0.9 },
  )
  assert.match(out, /left="0.5" right="0.9" top="0.75" bottom="0.75" header="0.3" footer="0.3"/)
})

test('withPageMargins inserts before the earliest successor when several are present', () => {
  const out = withPageMargins(
    '<worksheet><sheetData/><pageSetup/><tableParts count="1"/></worksheet>',
    { left: 1 },
  )
  assert.match(out, /<pageMargins [^>]*\/><pageSetup\/><tableParts/) // before pageSetup, the earliest
})

test('readPageSetup reads a portrait orientation', () => {
  assert.deepEqual(
    readPageSetup(
      encode('<worksheet><sheetData/><pageSetup orientation="portrait" scale="100"/></worksheet>'),
    ),
    { orientation: 'portrait', scale: 100 },
  )
})
