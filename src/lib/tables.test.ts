import assert from 'node:assert/strict'
import { test } from 'node:test'
import type { Container } from './container.js'
import { XlsxError } from './errors.js'
import { buildTablePart, extendTables, withTableParts } from './tables.js'

const encode = (text: string) => new TextEncoder().encode(text)

const SHEET_PATH = 'xl/worksheets/sheet1.xml'
const RELS_PATH = 'xl/worksheets/_rels/sheet1.xml.rels'

const sheetWith = (tableParts: string) => `<worksheet><sheetData/>${tableParts}</worksheet>`

const rels = (entries: string) => `<Relationships>${entries}</Relationships>`

const tableRel = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="${target}"/>`

const table = (ref: string, attributes = '', autoFilter = `<autoFilter ref="${ref}"/>`) =>
  `<table xmlns="http://x" id="1" name="T" displayName="T" ref="${ref}"${attributes}>${autoFilter}` +
  `<tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>`

function container(parts: Record<string, string>): Container {
  const map = new Map<string, Uint8Array>()
  for (const [path, content] of Object.entries(parts)) map.set(path, encode(content))
  return { parts: map }
}

/** One table over `ref`, wired to the sheet through its rels, plus overrides. */
function withOneTable(ref: string, extra: Record<string, string> = {}) {
  return container({
    [SHEET_PATH]: sheetWith('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'),
    [RELS_PATH]: rels(tableRel('rId1', '../tables/table1.xml')),
    'xl/tables/table1.xml': table(ref),
    ...extra,
  })
}

const extend = (parts: Container, written: string[]) =>
  extendTables(parts.parts.get(SHEET_PATH) ?? new Uint8Array(), SHEET_PATH, parts, written)

test('grows a table down to a cell written just below it', () => {
  const [result] = extend(withOneTable('A1:B2'), ['A3'])

  assert.equal(result?.path, 'xl/tables/table1.xml')
  assert.match(result?.xml ?? '', /<table[^>]*\sref="A1:B3"/)
  assert.match(result?.xml ?? '', /<autoFilter ref="A1:B3"\/>/)
})

test('does not grow a table across a gap', () => {
  assert.deepEqual(extend(withOneTable('A1:B2'), ['A5']), [])
})

test('grows a table with a huge column count without overflowing the call stack', () => {
  const columns = 200_000
  let cols = ''
  for (let index = 1; index <= columns; index++)
    cols += `<tableColumn id="${index}" name="c${index}"/>`
  const wide =
    '<table xmlns="http://x" id="1" name="T" displayName="T" ref="A1:C3"><autoFilter ref="A1:C3"/>' +
    `<tableColumns count="${columns}">${cols}</tableColumns></table>`
  const parts = withOneTable('A1:C3', { 'xl/tables/table1.xml': wide })

  const [result] = extend(parts, ['D1'])

  assert.match(result?.xml ?? '', new RegExp(`id="${columns + 1}" name="Column1"/>`))
})

test('a table rels part that is not valid utf-8 is a located unreadable-part', () => {
  const parts = new Map<string, Uint8Array>([
    [SHEET_PATH, encode(sheetWith('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'))],
    [RELS_PATH, new Uint8Array([0x80, 0x81, 0xff, 0xfe])],
    ['xl/tables/table1.xml', encode(table('A1:B2'))],
  ])
  const parts_: Container = { parts }

  assert.throws(
    () => extendTables(parts.get(SHEET_PATH) ?? new Uint8Array(), SHEET_PATH, parts_, ['A3']),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unreadable-part' && error.part === RELS_PATH,
  )
})

test('grows a table right to a cell written just past it, adding a column', () => {
  const [result] = extend(withOneTable('A1:B2'), ['C1'])

  assert.match(result?.xml ?? '', /<table[^>]*\sref="A1:C2"/)
  assert.match(result?.xml ?? '', /<autoFilter ref="A1:C2"\/>/)
  assert.match(result?.xml ?? '', /<tableColumns count="3">/)
  assert.match(result?.xml ?? '', /<tableColumn id="3" name="[^"]+"\/>/)
})

test('contiguous writes to the right add a column each', () => {
  const [result] = extend(withOneTable('A1:B2'), ['C1', 'D2'])

  assert.match(result?.xml ?? '', /ref="A1:D2"/)
  assert.match(result?.xml ?? '', /<tableColumns count="4">/)
})

test('does not grow right across a gap', () => {
  assert.deepEqual(extend(withOneTable('A1:B2'), ['D1']), [])
})

test('names an added column so it collides with none that exist', () => {
  const taken = table('A1:B2').replace('name="a"', 'name="Column1"').replace('name="b"', 'name="x"')
  const [result] = extend(withOneTable('A1:B2', { 'xl/tables/table1.xml': taken }), ['C1'])

  const names = [...(result?.xml ?? '').matchAll(/<tableColumn [^>]*name="([^"]+)"/g)].map(
    (m) => m[1],
  )
  assert.equal(names.length, 3)
  assert.equal(new Set(names).size, 3, `names must be unique: ${names.join(', ')}`)
})

test('resumes fresh column names past ones already taken when growing by several', () => {
  const named = table('A1:B2').replace('name="a"', 'name="Column2"').replace('name="b"', 'name="x"')
  const [result] = extend(withOneTable('A1:B2', { 'xl/tables/table1.xml': named }), ['C1', 'D1'])

  const names = [...(result?.xml ?? '').matchAll(/<tableColumn [^>]*name="([^"]+)"/g)].map(
    (m) => m[1],
  )
  assert.deepEqual(names, ['Column2', 'x', 'Column1', 'Column3'])
})

test('grows right even when existing columns declare no name or id', () => {
  const bare =
    '<table xmlns="http://x" ref="A1:B2"><autoFilter ref="A1:B2"/>' +
    '<tableColumns count="2"><tableColumn/><tableColumn/></tableColumns></table>'
  const [result] = extend(withOneTable('A1:B2', { 'xl/tables/table1.xml': bare }), ['C1'])

  assert.match(result?.xml ?? '', /ref="A1:C2"/)
  assert.match(result?.xml ?? '', /<tableColumn id="1" name="Column1"\/>/)
})

test('grows the ref right even for a table with no tableColumns element', () => {
  const noColumns = '<table xmlns="http://x" ref="A1:B2"><autoFilter ref="A1:B2"/></table>'
  const [result] = extend(withOneTable('A1:B2', { 'xl/tables/table1.xml': noColumns }), ['C1'])

  assert.match(result?.xml ?? '', /ref="A1:C2"/)
})

test('does not grow a table with a totals row', () => {
  const parts = withOneTable('A1:B3', {
    'xl/tables/table1.xml': table('A1:B3', ' totalsRowCount="1"'),
  })
  assert.deepEqual(extend(parts, ['A4']), [])
})

test('leaves a table part that is not valid utf-8 untouched', () => {
  const parts: Container = {
    parts: new Map<string, Uint8Array>([
      [
        SHEET_PATH,
        encode(sheetWith('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')),
      ],
      [RELS_PATH, encode(rels(tableRel('rId1', '../tables/table1.xml')))],
      ['xl/tables/table1.xml', new Uint8Array([0xff, 0xfe, 0x00])],
    ]),
  }

  assert.deepEqual(extend(parts, ['A3']), [])
})

test('leaves a table with a malformed ref untouched instead of crashing the write', () => {
  assert.deepEqual(extend(withOneTable('A1:B'), ['A3']), [])
})

test('leaves an autoFilter over a sub-range alone while growing the table', () => {
  const parts = withOneTable('A1:B2', {
    'xl/tables/table1.xml': table('A1:B2', '', '<autoFilter ref="A1:B1"/>'),
  })
  const [result] = extend(parts, ['A3'])

  assert.match(result?.xml ?? '', /<table[^>]*\sref="A1:B3"/)
  assert.match(result?.xml ?? '', /<autoFilter ref="A1:B1"\/>/)
})

test('does not grow one table into the rows of another', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith(
      '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>',
    ),
    [RELS_PATH]: rels(
      tableRel('rId1', '../tables/table1.xml') + tableRel('rId2', '../tables/table2.xml'),
    ),
    'xl/tables/table1.xml': table('A1:B2'),
    'xl/tables/table2.xml': table('A3:B4').replace('id="1"', 'id="2"'),
  })

  // A3 belongs to the second table, so the first must not swallow it.
  assert.deepEqual(
    extend(parts, ['A3']).filter((e) => e.path.endsWith('table1.xml')),
    [],
  )
})

test('grows a table without disturbing another that it does not reach', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith(
      '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>',
    ),
    [RELS_PATH]: rels(
      tableRel('rId1', '../tables/table1.xml') + tableRel('rId2', '../tables/table2.xml'),
    ),
    'xl/tables/table1.xml': table('A1:B2'),
    'xl/tables/table2.xml': table('A5:B6').replace('id="1"', 'id="2"'),
  })

  const grown = extend(parts, ['A3'])
  assert.equal(grown.length, 1)
  assert.match(grown[0]?.xml ?? '', /ref="A1:B3"/)
})

test('skips a table part whose id is missing', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith('<tableParts count="1"><tablePart/></tableParts>'),
    [RELS_PATH]: rels(tableRel('rId1', '../tables/table1.xml')),
    'xl/tables/table1.xml': table('A1:B2'),
  })
  assert.deepEqual(extend(parts, ['A3']), [])
})

test('skips a part that has no table element', () => {
  const parts = withOneTable('A1:B2', { 'xl/tables/table1.xml': '<notATable/>' })
  assert.deepEqual(extend(parts, ['A3']), [])
})

test('reads a table part id written without the r prefix', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith('<tableParts count="1"><tablePart id="rId1"/></tableParts>'),
    [RELS_PATH]: rels(tableRel('rId1', '../tables/table1.xml')),
    'xl/tables/table1.xml': table('A1:B2'),
  })
  assert.match(extend(parts, ['A3'])[0]?.xml ?? '', /ref="A1:B3"/)
})

test('does nothing for a sheet with no table parts', () => {
  const parts = container({ [SHEET_PATH]: '<worksheet><sheetData/></worksheet>' })
  assert.deepEqual(extend(parts, ['A3']), [])
})

test('does nothing when the sheet has no relationships part', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'),
  })
  assert.deepEqual(extend(parts, ['A3']), [])
})

test('skips a table part that is not in the package', () => {
  const parts = container({
    [SHEET_PATH]: sheetWith('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'),
    [RELS_PATH]: rels(tableRel('rId1', '../tables/gone.xml')),
  })
  assert.deepEqual(extend(parts, ['A3']), [])
})

test('skips a table whose ref is a single cell rather than a range', () => {
  const parts = withOneTable('A1:B2', {
    'xl/tables/table1.xml': table('A1'),
  })
  assert.deepEqual(extend(parts, ['A2']), [])
})

test('buildTablePart writes the ref, columns and style', () => {
  const xml = buildTablePart(1, {
    name: 'Sales',
    ref: 'A1:C5',
    columns: ['Name', 'Qty', 'Total'],
    style: 'TableStyleMedium2',
  })
  assert.match(xml, /<table [^>]*id="1" name="Sales" displayName="Sales" ref="A1:C5"/)
  assert.match(xml, /<autoFilter ref="A1:C5"\/>/)
  assert.match(
    xml,
    /<tableColumns count="3"><tableColumn id="1" name="Name"\/><tableColumn id="2" name="Qty"\/><tableColumn id="3" name="Total"\/>/,
  )
  assert.match(xml, /<tableStyleInfo name="TableStyleMedium2"/)
})

test('withTableParts opens a tableParts element and declares xmlns:r when missing', () => {
  const fresh = withTableParts(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    'rId1',
  )
  assert.match(
    fresh,
    /<worksheet [^>]*xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships"/,
  )
  assert.match(fresh, /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts><\/worksheet>/)
})

test('withTableParts joins an existing tableParts and bumps the count', () => {
  const joined = withTableParts(
    '<worksheet xmlns:r="r"><sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
    'rId2',
  )
  assert.match(joined, /<tableParts count="2"><tablePart r:id="rId1"\/><tablePart r:id="rId2"\/>/)
})

test('withTableParts places a fresh element before a worksheet extLst', () => {
  const out = withTableParts(
    '<worksheet xmlns:r="r"><sheetData/><extLst><ext/></extLst></worksheet>',
    'rId1',
  )
  assert.match(out, /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts><extLst>/)
})

test('withTableParts fills a self-closing tableParts element', () => {
  const out = withTableParts(
    '<worksheet xmlns:r="r"><sheetData/><tableParts count="1"/></worksheet>',
    'rId5',
  )
  assert.match(out, /<tableParts count="2"><tablePart r:id="rId5"\/><\/tableParts>/)
})

test('withTableParts appends to a rootless fragment without declaring xmlns:r', () => {
  assert.equal(
    withTableParts('<sheetData/>', 'rId1'),
    '<sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
  )
})

test('withTableParts joins a tableParts element that has no count attribute', () => {
  const out = withTableParts(
    '<worksheet xmlns:r="r"><sheetData/><tableParts><tablePart r:id="rId1"/></tableParts></worksheet>',
    'rId2',
  )
  assert.match(out, /<tableParts count="1"><tablePart r:id="rId1"\/><tablePart r:id="rId2"\/>/)
})
