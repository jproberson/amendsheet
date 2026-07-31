import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  readHeaderFooter,
  readPageBreaks,
  readPageMargins,
  readPageSetup,
  withColumnBreaks,
  withHeaderFooter,
  withPageMargins,
  withPageSetup,
  withRowBreaks,
} from './page.js'

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

test('withHeaderFooter inserts a fresh element after pageSetup and before its successors', () => {
  const out = withHeaderFooter(
    '<worksheet><sheetData/><pageMargins/><pageSetup/><tableParts count="1"/></worksheet>',
    { header: { center: 'Report' }, footer: { right: '&P of &N' } },
  )
  assert.match(
    out,
    /<pageSetup\/><headerFooter><oddHeader>&amp;CReport<\/oddHeader><oddFooter>&amp;R&amp;P of &amp;N<\/oddFooter><\/headerFooter><tableParts/,
  )
})

test('withHeaderFooter builds only the sections given, in left-center-right order', () => {
  const out = withHeaderFooter('<worksheet><sheetData/></worksheet>', {
    header: { right: 'R', left: 'L' },
  })
  assert.match(out, /<oddHeader>&amp;LL&amp;RR<\/oddHeader>/)
})

test('withHeaderFooter replaces the oddHeader within an existing element, keeping siblings and attributes', () => {
  const out = withHeaderFooter(
    '<worksheet><sheetData/><headerFooter differentFirst="1"><oddHeader>&amp;COld</oddHeader><oddFooter>&amp;CFoot</oddFooter><firstHeader>&amp;CFirst</firstHeader></headerFooter></worksheet>',
    { header: { center: 'New' } },
  )
  assert.match(
    out,
    /<headerFooter differentFirst="1"><oddHeader>&amp;CNew<\/oddHeader><oddFooter>&amp;CFoot<\/oddFooter><firstHeader>&amp;CFirst<\/firstHeader><\/headerFooter>/,
  )
})

test('withHeaderFooter inserts an oddFooter after the oddHeader to keep child order', () => {
  const out = withHeaderFooter(
    '<worksheet><sheetData/><headerFooter><oddHeader>&amp;CH</oddHeader></headerFooter></worksheet>',
    { footer: { center: 'F' } },
  )
  assert.match(
    out,
    /<headerFooter><oddHeader>&amp;CH<\/oddHeader><oddFooter>&amp;CF<\/oddFooter><\/headerFooter>/,
  )
})

test('withHeaderFooter xml-escapes the section text', () => {
  const out = withHeaderFooter('<worksheet><sheetData/></worksheet>', {
    header: { center: 'a<b&c' },
  })
  assert.match(out, /<oddHeader>&amp;Ca&lt;b&amp;c<\/oddHeader>/)
})

test('readHeaderFooter parses the odd header and footer into sections', () => {
  const xml =
    '<worksheet><sheetData/><headerFooter><oddHeader>&amp;LLeft&amp;CMid&amp;RRight</oddHeader><oddFooter>&amp;CPage &amp;P</oddFooter></headerFooter></worksheet>'
  assert.deepEqual(readHeaderFooter(encode(xml)), {
    header: { left: 'Left', center: 'Mid', right: 'Right' },
    footer: { center: 'Page &P' },
  })
})

test('readHeaderFooter puts uncoded leading text in the center section', () => {
  assert.deepEqual(
    readHeaderFooter(
      encode(
        '<worksheet><sheetData/><headerFooter><oddHeader>Plain</oddHeader></headerFooter></worksheet>',
      ),
    ),
    { header: { center: 'Plain' } },
  )
})

test('readHeaderFooter treats a doubled ampersand as a literal, not a section switch', () => {
  const xml =
    '<worksheet><sheetData/><headerFooter><oddHeader>&amp;CQ&amp;&amp;LA</oddHeader></headerFooter></worksheet>'
  assert.deepEqual(readHeaderFooter(encode(xml)), { header: { center: 'Q&&LA' } })
})

test('readHeaderFooter returns nothing when there is no headerFooter', () => {
  assert.deepEqual(readHeaderFooter(encode('<worksheet><sheetData/></worksheet>')), {})
})

test('a header and footer written by withHeaderFooter read back through readHeaderFooter', () => {
  const out = withHeaderFooter('<worksheet><sheetData/></worksheet>', {
    header: { center: 'Q&&A Report' },
    footer: { left: 'Confidential', right: '&P of &N' },
  })
  assert.deepEqual(readHeaderFooter(encode(out)), {
    header: { center: 'Q&&A Report' },
    footer: { left: 'Confidential', right: '&P of &N' },
  })
})

test('withRowBreaks inserts a fresh rowBreaks with a manual break after pageSetup', () => {
  const out = withRowBreaks('<worksheet><sheetData/><pageSetup/></worksheet>', [9])
  assert.match(
    out,
    /<pageSetup\/><rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"\/><\/rowBreaks><\/worksheet>/,
  )
})

test('withColumnBreaks inserts a fresh colBreaks that lands after an existing rowBreaks', () => {
  const out = withColumnBreaks(
    '<worksheet><sheetData/><rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks></worksheet>',
    [3],
  )
  assert.match(
    out,
    /<\/rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"\/><\/colBreaks>/,
  )
})

test('withRowBreaks merges into an existing container, sorting, deduping and recounting', () => {
  const out = withRowBreaks(
    '<worksheet><sheetData/><rowBreaks count="1" manualBreakCount="1"><brk id="20" max="16383" man="1"/></rowBreaks></worksheet>',
    [9, 20],
  )
  assert.match(
    out,
    /<rowBreaks count="2" manualBreakCount="2"><brk id="9" max="16383" man="1"\/><brk id="20" max="16383" man="1"\/><\/rowBreaks>/,
  )
})

test('withRowBreaks keeps an existing automatic break and counts the manual ones apart', () => {
  const out = withRowBreaks(
    '<worksheet><sheetData/><rowBreaks count="1" manualBreakCount="0"><brk id="30" max="16383"/></rowBreaks></worksheet>',
    [9],
  )
  assert.match(
    out,
    /<rowBreaks count="2" manualBreakCount="1"><brk id="9" max="16383" man="1"\/><brk id="30" max="16383"\/><\/rowBreaks>/,
  )
})

test('withRowBreaks replaces an empty self-closing container', () => {
  const out = withRowBreaks('<worksheet><sheetData/><rowBreaks/></worksheet>', [9])
  assert.match(
    out,
    /<rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"\/><\/rowBreaks>/,
  )
})

test('readPageBreaks returns rows as 1-based numbers and columns as letters, sorted', () => {
  const xml =
    '<worksheet><sheetData/><rowBreaks count="2" manualBreakCount="2"><brk id="20" max="16383" man="1"/><brk id="9" max="16383" man="1"/></rowBreaks><colBreaks count="2" manualBreakCount="2"><brk id="7" max="1048575" man="1"/><brk id="3" max="1048575" man="1"/></colBreaks></worksheet>'
  assert.deepEqual(readPageBreaks(encode(xml)), { rows: [10, 21], columns: ['D', 'H'] })
})

test('readPageBreaks returns empty arrays when there are none', () => {
  assert.deepEqual(readPageBreaks(encode('<worksheet><sheetData/></worksheet>')), {
    rows: [],
    columns: [],
  })
})

test('readPageBreaks ignores an empty self-closing container', () => {
  assert.deepEqual(
    readPageBreaks(
      encode('<worksheet><sheetData/><rowBreaks count="0" manualBreakCount="0"/></worksheet>'),
    ),
    { rows: [], columns: [] },
  )
})

test('readPageBreaks treats a break with no id as the sheet top', () => {
  const xml =
    '<worksheet><sheetData/><rowBreaks count="1" manualBreakCount="1"><brk max="16383" man="1"/></rowBreaks></worksheet>'
  assert.deepEqual(readPageBreaks(encode(xml)), { rows: [1], columns: [] })
})

test('breaks written by withRowBreaks and withColumnBreaks read back through readPageBreaks', () => {
  let xml = '<worksheet><sheetData/></worksheet>'
  xml = withRowBreaks(xml, [9, 20])
  xml = withColumnBreaks(xml, [3])
  assert.deepEqual(readPageBreaks(encode(xml)), { rows: [10, 21], columns: ['D'] })
})
