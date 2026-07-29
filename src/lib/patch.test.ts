import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  type CellInput,
  type SheetEdits,
  checkProtection,
  checkWritable,
  indexSheet,
  mergeRangeReference,
  patchSheet as patchSheetBytes,
  readColumnGroupLevels,
  readColumnWidths,
  readConditionalFormats,
  readDataValidations,
  readHiddenColumns,
  readHiddenRows,
  readRowGroupLevels,
  readRowHeights,
  readSheetProtection,
  readSheetView,
} from './patch.js'
import { XlsxError } from './errors.js'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes)
const patchSheet = (
  source: string,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
  sharedStrings?: ReadonlyMap<string, number>,
  styleOverrides?: ReadonlyMap<string, number>,
  sheet?: SheetEdits,
) =>
  decode(patchSheetBytes(encode(source), edits, date1904, sharedStrings, styleOverrides, {}, sheet))

const sheet = (rows: string) =>
  `<?xml version="1.0"?><worksheet xmlns="http://x"><cols><col min="1" max="1" width="20"/></cols>` +
  `<sheetData>${rows}</sheetData>` +
  `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><pageMargins left="0.7"/></worksheet>`

const ROWS =
  '<row r="1" ht="30" customHeight="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="4"><v>1</v></c></row>' +
  '<row r="2" hidden="1"><c r="A2"><v>2</v></c></row>'

const patch = (edits: Array<[string, CellInput]>) => patchSheet(sheet(ROWS), new Map(edits), false)

test('returns the source untouched when there is nothing to do', () => {
  const source = sheet(ROWS)

  assert.equal(patchSheet(source, new Map(), false), source)
})

test('writes sheet protection after sheetData and before mergeCells', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    protection: {},
  })

  assert.match(
    patched,
    /<\/sheetData><sheetProtection sheet="1" objects="1" scenarios="1"\/><mergeCells/,
  )
})

test('replaces a sheet protection the sheet already declares', () => {
  const source =
    '<worksheet xmlns="http://x"><sheetData/><sheetProtection sheet="1" formatCells="1"/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    protection: { formatCells: true },
  })

  assert.doesNotMatch(patched, /formatCells="1"/)
  assert.match(patched, /<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"\/>/)
  assert.equal((patched.match(/<sheetProtection/g) ?? []).length, 1)
})

test('writes sheet protection with the document namespace prefix', () => {
  const source = '<x:worksheet xmlns:x="http://x"><x:sheetData/></x:worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, { protection: {} })

  assert.match(patched, /<x:sheetData\/><x:sheetProtection sheet="1"/)
})

test('removing sheet protection drops the element the sheet had', () => {
  const source =
    '<worksheet xmlns="http://x"><sheetData/><sheetProtection sheet="1"/><pageMargins left="0.7"/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    protection: 'remove',
  })

  assert.doesNotMatch(patched, /sheetProtection/)
  assert.match(patched, /<sheetData\/><pageMargins/)
})

test('removing sheet protection a sheet lacks changes nothing', () => {
  const source = sheet(ROWS)

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    protection: 'remove',
  })

  assert.equal(patched, source)
})

test('adds a mergeCell to the mergeCells the sheet already has', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    merges: ['C1:D2'],
  })

  assert.match(patched, /<mergeCells count="2"><mergeCell ref="A1:B1"\/><mergeCell ref="C1:D2"\/>/)
})

test('opens a mergeCells element when the sheet has none', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/><pageMargins left="0.7"/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, { merges: ['A1:B2'] })

  assert.match(
    patched,
    /<sheetData\/><mergeCells count="1"><mergeCell ref="A1:B2"\/><\/mergeCells>/,
  )
})

test('adds an autoFilter after sheetData and before mergeCells', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    autoFilter: 'A1:B1',
  })

  assert.match(patched, /<\/sheetData><autoFilter ref="A1:B1"\/><mergeCells/)
})

test('replaces an autoFilter the sheet already has, filters and all', () => {
  const selfClosing =
    '<worksheet xmlns="http://x"><sheetData/><autoFilter ref="A1:C1"/></worksheet>'
  const withFilters =
    '<worksheet xmlns="http://x"><sheetData/><autoFilter ref="A1:C1"><filterColumn colId="0"/></autoFilter></worksheet>'

  for (const source of [selfClosing, withFilters]) {
    const patched = patchSheet(source, new Map(), false, undefined, undefined, {
      autoFilter: 'A1:D1',
    })
    assert.match(patched, /<autoFilter ref="A1:D1"\/>/)
    assert.equal((patched.match(/<autoFilter/g) ?? []).length, 1)
    assert.doesNotMatch(patched, /filterColumn/)
  }
})

test('freeze adds a pane before cols and preserves the sheetView', () => {
  const fresh = patchSheet(
    '<worksheet xmlns="http://x"><cols><col min="1" max="1" width="5"/></cols><sheetData/></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { freeze: 'B2' },
  )
  assert.match(
    fresh,
    /<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/><\/sheetView><\/sheetViews><cols>/,
  )

  const kept = patchSheet(
    '<worksheet xmlns="http://x"><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1"/></sheetView></sheetViews><sheetData/></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { freeze: 'B2' },
  )
  assert.match(kept, /<sheetView workbookViewId="0"><pane[^>]*\/><selection/)

  const replaced = patchSheet(
    '<worksheet xmlns="http://x"><sheetViews><sheetView workbookViewId="0"><pane xSplit="5" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { freeze: 'B2' },
  )
  assert.equal((replaced.match(/<pane/g) ?? []).length, 1)
  assert.match(replaced, /topLeftCell="B2"/)
})

test('freeze takes the split from the cell, rows or columns only', () => {
  const at = (cell: string) =>
    patchSheet(
      '<worksheet xmlns="http://x"><sheetData/></worksheet>',
      new Map(),
      false,
      undefined,
      undefined,
      {
        freeze: cell,
      },
    )
  assert.match(at('A2'), /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft"/)
  assert.match(at('B1'), /<pane xSplit="1" topLeftCell="B1" activePane="topRight"/)
  assert.match(at('B2'), /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight"/)
})

test('hides rows and columns, composing with any width or height', () => {
  const rows = patchSheet(
    '<worksheet xmlns="http://x"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { hiddenRows: new Set([1, 2]), rowHeights: new Map([[1, 15]]) },
  )
  assert.match(rows, /<row hidden="1" customHeight="1" ht="15" r="1">/)
  assert.match(rows, /<row r="2" hidden="1"><\/row>/)

  const cols = patchSheet(
    '<worksheet xmlns="http://x"><cols><col min="1" max="5" width="20"/></cols><sheetData/></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { hiddenColumns: new Set([3]) },
  )
  assert.match(
    cols,
    /<col min="1" max="2" width="20"\/><col hidden="1" min="3" max="3" width="20"\/><col min="4" max="5" width="20"\/>/,
  )

  const both = patchSheet(
    '<worksheet xmlns="http://x"><sheetData/></worksheet>',
    new Map(),
    false,
    undefined,
    undefined,
    { columnWidths: new Map([[5, 20]]), hiddenColumns: new Set([5]) },
  )
  assert.match(both, /<col min="5" max="5" width="20" customWidth="1" hidden="1"\/>/)
})

test('a merge the sheet already declares is not added twice', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    merges: ['A1:B1'],
  })

  assert.equal(patched, decode(encode(sheet(ROWS))))
})

test('opens a self closing mergeCells element to hold a new merge', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/><mergeCells/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, { merges: ['A1:B2'] })

  assert.match(patched, /<mergeCells count="1"><mergeCell ref="A1:B2"\/><\/mergeCells>/)
})

test('a new merge lands after sheet protection added in the same write', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    protection: {},
    merges: ['A1:B2'],
  })

  assert.match(patched, /<sheetProtection[^>]*\/><mergeCells count="1">/)
})

test('a new mergeCells opens after a sheet protection the file already had', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/><sheetProtection sheet="1"/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, { merges: ['A1:B2'] })

  assert.match(
    patched,
    /<sheetProtection sheet="1"\/><mergeCells count="1"><mergeCell ref="A1:B2"\/>/,
  )
})

test('sets the height on an existing row, keeping its other attributes', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    rowHeights: new Map([[2, 40]]),
  })

  assert.match(patched, /<row customHeight="1" ht="40" r="2" hidden="1">/)
})

test('replaces a height a row already declares', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    rowHeights: new Map([[1, 15]]),
  })

  assert.match(patched, /<row r="1" ht="15" customHeight="1">/)
  assert.doesNotMatch(patched, /ht="30"/)
})

test('adds a new empty row for a height on a row that has none', () => {
  const source =
    '<worksheet xmlns="http://x"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    rowHeights: new Map([[3, 22]]),
  })

  assert.match(patched, /<row r="3" ht="22" customHeight="1"><\/row>/)
})

test('a height and a value on the same new row land together', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/></worksheet>'

  const patched = patchSheet(
    source,
    new Map<string, CellInput>([['A5', 7]]),
    false,
    undefined,
    undefined,
    { rowHeights: new Map([[5, 33]]) },
  )

  assert.match(patched, /<row r="5" ht="33" customHeight="1"><c r="A5"><v>7<\/v><\/c><\/row>/)
})

test('a height on a self closing row that also gains a cell keeps both', () => {
  const source = '<worksheet xmlns="http://x"><sheetData><row r="4"/></sheetData></worksheet>'

  const patched = patchSheet(
    source,
    new Map<string, CellInput>([['B4', 9]]),
    false,
    undefined,
    undefined,
    { rowHeights: new Map([[4, 18]]) },
  )

  assert.match(patched, /<row customHeight="1" ht="18" r="4"><c r="B4"><v>9<\/v><\/c><\/row>/)
})

test('opens a cols element before sheetData for a new column width', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    columnWidths: new Map([[3, 18]]),
  })

  assert.match(
    patched,
    /<cols><col min="3" max="3" width="18" customWidth="1"\/><\/cols><sheetData\/>/,
  )
})

test('replaces the width of a single-column col', () => {
  // sheet(ROWS) has <col min="1" max="1" width="20"/>
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    columnWidths: new Map([[1, 42]]),
  })

  assert.match(patched, /<col customWidth="1" min="1" max="1" width="42"\/>/)
  assert.doesNotMatch(patched, /width="20"/)
})

test('splits a multi-column col so the rest keeps its width', () => {
  const source =
    '<worksheet xmlns="http://x"><cols><col min="1" max="5" width="10"/></cols><sheetData/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    columnWidths: new Map([[3, 30]]),
  })

  assert.match(
    patched,
    /<col min="1" max="2" width="10"\/><col customWidth="1" min="3" max="3" width="30"\/><col min="4" max="5" width="10"\/>/,
  )
})

test('appends a column width into a cols element that has one', () => {
  const patched = patchSheet(sheet(ROWS), new Map(), false, undefined, undefined, {
    columnWidths: new Map([[4, 25]]),
  })

  assert.match(
    patched,
    /<col min="1" max="1" width="20"\/><col min="4" max="4" width="25" customWidth="1"\/>/,
  )
})

test('splits a col for two widths within it, in column order', () => {
  const source =
    '<worksheet xmlns="http://x"><cols><col min="1" max="5" width="10"/></cols><sheetData/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    columnWidths: new Map([
      [3, 30],
      [2, 15],
    ]),
  })

  assert.match(
    patched,
    /<col min="1" max="1" width="10"\/><col customWidth="1" min="2" max="2" width="15"\/><col customWidth="1" min="3" max="3" width="30"\/><col min="4" max="5" width="10"\/>/,
  )
})

test('appends two new column widths in column order', () => {
  const source = '<worksheet xmlns="http://x"><sheetData/></worksheet>'

  const patched = patchSheet(source, new Map(), false, undefined, undefined, {
    columnWidths: new Map([
      [5, 10],
      [3, 20],
    ]),
  })

  assert.match(
    patched,
    /<cols><col min="3" max="3" width="20" customWidth="1"\/><col min="5" max="5" width="10" customWidth="1"\/><\/cols>/,
  )
})

test('reads sheet protection back as the permissions it grants', () => {
  const bytes = encode(
    '<worksheet xmlns="http://x"><sheetData/>' +
      '<sheetProtection sheet="1" objects="1" scenarios="0" formatCells="0" selectLockedCells="1"/>' +
      '</worksheet>',
  )

  assert.deepEqual(readSheetProtection(bytes), {
    editObjects: false,
    editScenarios: true,
    formatCells: true,
    selectLockedCells: false,
  })
})

test('reads no protection when the element is absent or turned off', () => {
  assert.equal(
    readSheetProtection(encode('<worksheet xmlns="http://x"><sheetData/></worksheet>')),
    undefined,
  )
  assert.equal(
    readSheetProtection(
      encode('<worksheet xmlns="http://x"><sheetData/><sheetProtection sheet="0"/></worksheet>'),
    ),
    undefined,
  )
})

test('restyling a cell rewrites its style and keeps its value and formula', () => {
  const source = sheet('<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row>')

  const patched = patchSheet(source, new Map(), false, undefined, new Map([['A1', 7]]))

  // s is inserted after the element name; attribute order is insignificant.
  assert.match(patched, /<c s="7" r="A1"><f>SUM\(B1:B2\)<\/f><v>3<\/v><\/c>/)
})

test('restyling a cell that carries a style replaces the style', () => {
  const source = sheet('<row r="1"><c r="A1" s="2"><v>3</v></c></row>')

  const patched = patchSheet(source, new Map(), false, undefined, new Map([['A1', 9]]))

  assert.match(patched, /<c r="A1" s="9"><v>3<\/v><\/c>/)
})

test('restyling a cell that is not there yet adds an empty styled cell', () => {
  const source = sheet('<row r="1"><c r="A1"><v>1</v></c></row>')

  const patched = patchSheet(source, new Map(), false, undefined, new Map([['B1', 4]]))

  assert.match(patched, /<c r="A1"><v>1<\/v><\/c><c r="B1" s="4"\/><\/row>/)
})

test('a value write and a restyle of another cell both land', () => {
  const source = sheet('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>')

  const patched = patchSheet(
    source,
    new Map<string, CellInput>([['A1', 9]]),
    false,
    undefined,
    new Map([
      ['A1', 3],
      ['B1', 5],
    ]),
  )

  assert.match(patched, /<c r="A1" s="3"><v>9<\/v><\/c>/)
  assert.match(patched, /<c s="5" r="B1"><v>2<\/v><\/c>/)
})

test('a styled cell no column letter can name does not break the index', () => {
  // XFE is column 16385, past the last a sheet can hold; the reader accepts it
  // leniently, so indexing must skip it rather than throw and take set() down.
  const source = sheet(
    '<row r="1"><c r="XFE1" s="1"><v>1</v></c><c r="A1" s="2"><v>2</v></c></row>',
  )

  const index = indexSheet(new TextEncoder().encode(source))

  assert.equal(index.styles.get('A1'), 2)
})

test('replaces the value of a cell', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<c r="A2"><v>99<\/v><\/c>/)
})

test('keeps the style of the cell it replaces', () => {
  const patched = patch([['B1', 7]])

  assert.match(patched, /<c r="B1" s="4"><v>7<\/v><\/c>/)
})

test('leaves every other cell alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<c r="A1" t="s"><v>0<\/v><\/c>/)
  assert.match(patched, /<c r="B1" s="4"><v>1<\/v><\/c>/)
})

test('leaves row attributes alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<row r="1" ht="30" customHeight="1">/)
  assert.match(patched, /<row r="2" hidden="1">/)
})

test('leaves everything outside the cell data alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<col min="1" max="1" width="20"\/>/)
  assert.match(patched, /<mergeCell ref="A1:B1"\/>/)
  assert.match(patched, /<pageMargins left="0.7"\/>/)
})

test('replaces a cell that was written self closing', () => {
  const source = sheet('<row r="1"><c r="A1" s="2"/></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['A1', 5]]), false)

  assert.match(patched, /<c r="A1" s="2"><v>5<\/v><\/c>/)
})

test('writes text as an inline string', () => {
  const patched = patch([['A2', 'hello']])

  assert.match(patched, /<c r="A2" t="inlineStr"><is><t>hello<\/t><\/is><\/c>/)
})

test('escapes text that would break the markup', () => {
  const patched = patch([['A2', 'a & b < c']])

  assert.match(patched, /<t>a &amp; b &lt; c<\/t>/)
})

test('leaves quotes alone in text, which need no escaping', () => {
  const patched = patch([['A2', 'say "hello"']])

  assert.match(patched, /<t>say "hello"<\/t>/)
})

test('writes a boolean', () => {
  const patched = patch([['A2', true]])

  assert.match(patched, /<c r="A2" t="b"><v>1<\/v><\/c>/)
})

test('writes null as a cell with no value', () => {
  const patched = patch([['B1', null]])

  assert.match(patched, /<c r="B1" s="4"\/>/)
})

test('writes a date as its serial', () => {
  const patched = patch([['A2', new Date(2024, 0, 1)]])

  assert.match(patched, /<c r="A2"><v>45292<\/v><\/c>/)
})

test('writes a date against the workbook epoch', () => {
  const patched = patchSheet(
    sheet(ROWS),
    new Map<string, CellInput>([['A2', new Date(1904, 0, 1)]]),
    true,
  )

  assert.match(patched, /<c r="A2"><v>0<\/v><\/c>/)
})

test('applies several edits at once', () => {
  const patched = patch([
    ['A2', 1],
    ['B1', 2],
  ])

  assert.match(patched, /<c r="B1" s="4"><v>2<\/v><\/c>/)
  assert.match(patched, /<c r="A2"><v>1<\/v><\/c>/)
})

test('adds a cell that was not in the sheet yet', () => {
  const patched = patch([['Z9', 1]])

  assert.match(patched, /<row r="9"><c r="Z9"><v>1<\/v><\/c><\/row>/)
})

test('rejects a number that cannot be written', () => {
  assert.throws(() => patch([['A2', Number.POSITIVE_INFINITY]]), /cannot hold/)
})

test('keeps whitespace on an inline string', () => {
  const patched = patch([['A2', '  padded  ']])

  assert.match(patched, /<t xml:space="preserve"> {2}padded {2}<\/t>/)
})

test('rejects text that xml cannot represent', () => {
  assert.throws(() => patch([['A2', `a${String.fromCharCode(7)}b`]]), /cannot be written to xml/i)
})

test('refuses to overwrite the master of a shared formula', () => {
  const source = sheet(
    '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>',
  )

  assert.throws(
    () => patchSheet(source, new Map<string, CellInput>([['A1', 9]]), false),
    /shared formula/i,
  )
})

test('a malformed cell reference on the edit path is a located file fault', () => {
  // The caller's edit is fine; the sheet's own <c r="A"> is not. It must surface
  // as located invalid content, not a caller-fault bad-reference or a raw crash.
  const source = sheet('<row r="1"><c r="A"><v>1</v></c></row>')
  const at = { sheet: 'Sheet1', part: 'xl/worksheets/sheet1.xml' }

  assert.throws(
    () =>
      patchSheetBytes(
        encode(source),
        new Map<string, CellInput>([['B2', 5]]),
        false,
        undefined,
        undefined,
        at,
      ),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'invalid-content' &&
      error.reference === 'A' &&
      error.sheet === 'Sheet1',
  )
})

test('a row whose r is not a number on the edit path is a located file fault', () => {
  const source = sheet('<row r="abc"><c r="A1"><v>1</v></c></row>')

  assert.throws(
    () => patchSheet(source, new Map<string, CellInput>([['B2', 5]]), false),
    (error: unknown) => error instanceof XlsxError && error.code === 'invalid-content',
  )
})

test('writes over a cell that follows a shared formula', () => {
  const source = sheet(
    '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>',
  )

  const patched = patchSheet(source, new Map<string, CellInput>([['A2', 9]]), false)

  assert.match(patched, /<c r="A2"><v>9<\/v><\/c>/)
})

test('writes over an ordinary formula cell', () => {
  const source = sheet('<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['A1', 9]]), false)

  assert.match(patched, /<c r="A1"><v>9<\/v><\/c>/)
})

test('recognises a shared formula master written self closing', () => {
  const source = sheet('<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0"/><v>2</v></c></row>')

  assert.throws(
    () => patchSheet(source, new Map<string, CellInput>([['A1', 9]]), false),
    /shared formula/i,
  )
})

test('writes over a dependent written with a closing tag', () => {
  const source = sheet(
    '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" si="0"></f><v>4</v></c></row>',
  )

  const patched = patchSheet(source, new Map<string, CellInput>([['A2', 9]]), false)

  assert.match(patched, /<c r="A2"><v>9<\/v><\/c>/)
})

test('widens a dimension written with a closing tag', () => {
  const source =
    '<worksheet><dimension ref="A1:B2"></dimension><sheetData>' +
    '<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

  const patched = patchSheet(source, new Map<string, CellInput>([['D5', 4]]), false)

  assert.equal(patched.includes('</dimension>'), false, 'a dangling close tag was left behind')
  assert.match(patched, /<dimension ref="A1:D5"\/>/)
})

test('writes a formula', () => {
  const patched = patch([['A2', { formula: 'SUM(B1:B9)' }]])

  assert.match(patched, /<c r="A2"><f>SUM\(B1:B9\)<\/f><\/c>/)
})

test('does not write a cached result, so it is recalculated', () => {
  const patched = patch([['A2', { formula: 'SUM(B1:B9)' }]])

  assert.equal(/<c r="A2">.*?<v>/.test(patched), false)
})

test('escapes a formula that would break the markup', () => {
  const patched = patch([['A2', { formula: 'IF(B1<C1,"a & b","c")' }]])

  assert.match(patched, /<f>IF\(B1&lt;C1,"a &amp; b","c"\)<\/f>/)
})

test('keeps the style of the cell a formula replaces', () => {
  const patched = patch([['B1', { formula: 'A1*2' }]])

  assert.match(patched, /<c r="B1" s="4"><f>A1\*2<\/f><\/c>/)
})

test('leaves text that merely starts with an equals sign as text', () => {
  const patched = patch([['A2', '=not a formula']])

  assert.match(patched, /t="inlineStr"><is><t>=not a formula<\/t>/)
})

test('refuses a formula that xml cannot hold', () => {
  assert.throws(
    () => patch([['A2', { formula: `SUM(${String.fromCharCode(7)})` }]]),
    /cannot be written to xml/i,
  )
})

test('writes a formula into a cell that is not there yet', () => {
  const patched = patch([['Z9', { formula: 'A1+1' }]])

  assert.match(patched, /<row r="9"><c r="Z9"><f>A1\+1<\/f><\/c><\/row>/)
})

test('places new rows against a sheet whose rows are not in order', () => {
  const patched = patchSheet(
    sheet('<row r="9"><c r="A9"><v>9</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'),
    new Map<string, CellInput>([
      ['A5', 5],
      ['A7', 7],
      ['A20', 20],
    ]),
    false,
  )

  assert.match(
    patched,
    /<sheetData><row r="5">.*<row r="7">.*<row r="9">.*<row r="2">.*<row r="20">.*<\/sheetData>/,
  )
})

test('adds several new rows in ascending order', () => {
  const patched = patch([
    ['A9', 9],
    ['A5', 5],
    ['A7', 7],
  ])

  assert.match(patched, /<row r="5">.*<row r="7">.*<row r="9">/)
})

test('escapes a carriage return so a parser cannot normalise it away', () => {
  // XML 1.0 requires a parser to turn a literal CR or CRLF in content into a
  // bare LF before the application sees it, so writing one raw loses it.
  const patched = patch([['A2', 'a\r\nb']])

  assert.equal(patched.includes('\r'), false, 'no raw carriage return may reach the file')
  assert.match(patched, /a&#13;\nb/)
})

test('numbers a row without an r the way a reader counting rows does', () => {
  // Excel writes r on every row, but the format leaves it optional, and a row
  // that omits it is the one after the row before it. Counting rows seen so far
  // instead put this cell in row 2, so cells() and set() disagreed about which
  // row it was, and an edit to A6 appended a second row 6.
  const source = sheet('<row r="5"><c r="A5"><v>1</v></c></row><row><c r="A6"><v>2</v></c></row>')

  const patched = patchSheet(source, new Map([['A6', 99]]), false)

  assert.match(patched, /<c r="A6"><v>99<\/v><\/c>/)
  assert.equal(patched.match(/<row/g)?.length, 2, 'a second row 6 was appended')
})

test('refuses to overwrite the cell an array formula spills from', () => {
  // The other cells in the range hold cached values and no formula, so
  // replacing the anchor leaves them owned by nothing.
  const spilling = sheet(
    '<row r="1"><c r="A1"><f t="array" ref="A1:B2">TRANSPOSE(D1:E2)</f><v>1</v></c></row>',
  )

  assert.throws(
    () => patchSheet(spilling, new Map<string, CellInput>([['A1', 5]]), false),
    /A1:B2|array/i,
  )
})

test('refuses to overwrite the cell a data table spills from', () => {
  const table = sheet('<row r="1"><c r="A1"><f t="dataTable" ref="A1:A3"/><v>1</v></c></row>')

  assert.throws(
    () => patchSheet(table, new Map<string, CellInput>([['A1', 5]]), false),
    /A1:A3|table/i,
  )
})

test('allows writing a cell inside an array range that is not the anchor', () => {
  const spilling = sheet(
    '<row r="1"><c r="A1"><f t="array" ref="A1:B2">TRANSPOSE(D1:E2)</f><v>1</v></c>' +
      '<c r="B1"><v>2</v></c></row>',
  )

  assert.match(patchSheet(spilling, new Map<string, CellInput>([['B1', 5]]), false), /<v>5<\/v>/)
})

test('refuses an object that is not a date and names no formula', () => {
  // A JS caller, a JSON payload or an any at a boundary reaches here, and
  // reading .formula off it threw a TypeError, outside the error contract.
  for (const value of [{}, { formula: 5 }, []]) {
    assert.throws(
      () => checkWritable('A1', value, false),
      (error: unknown) =>
        error instanceof XlsxError && error.code === 'unwritable-value' && error.reference === 'A1',
      `${JSON.stringify(value)} is not a cell value`,
    )
  }
})

test('mergeRangeReference refuses a non-string range', () => {
  const refuses = (range: unknown) =>
    assert.throws(
      () => mergeRangeReference(range),
      (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
    )
  refuses(123)
  refuses(null)
  refuses(undefined)
})

test('checkProtection refuses an argument that is not an options object', () => {
  const refuses = (options: unknown) =>
    assert.throws(
      () => checkProtection(options),
      (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
    )
  refuses(null)
  refuses('locked')
  refuses(7)
  checkProtection({}) // a valid empty options object does not throw
  checkProtection({ formatCells: true })
})

test('readRowHeights reads ht and skips a row without one or with a bad value', () => {
  const bytes = encode(
    '<worksheet><sheetData>' +
      '<row r="1" ht="30" customHeight="1"/><row r="2"/><row r="3" ht="x"/>' +
      '</sheetData></worksheet>',
  )
  const heights = readRowHeights(bytes)
  assert.equal(heights.get(1), 30)
  assert.equal(heights.has(2), false)
  assert.equal(heights.has(3), false)
})

test('readColumnWidths reads width ranges and skips a col without a width or with a bad one', () => {
  const bytes = encode(
    '<worksheet><cols>' +
      '<col min="1" max="3" width="10" customWidth="1"/><col min="5" max="5"/><col min="7" max="x" width="9"/>' +
      '</cols><sheetData/></worksheet>',
  )
  assert.deepEqual(readColumnWidths(bytes), [{ min: 1, max: 3, width: 10 }])
})

test('readSheetView reads gridlines, headings, zoom, freeze, tab colour and autoFilter', () => {
  const bytes = encode(
    '<worksheet><sheetPr><tabColor rgb="FFFF0000"/></sheetPr>' +
      '<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" zoomScale="150">' +
      '<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>' +
      '<sheetData/><autoFilter ref="A1:C1"/></worksheet>',
  )
  const view = readSheetView(bytes)
  assert.equal(view.gridlines, false)
  assert.equal(view.headings, false)
  assert.equal(view.zoom, 150)
  assert.equal(view.frozen, 'B2')
  assert.equal(view.tabColor, 'FFFF0000')
  assert.equal(view.autoFilter, 'A1:C1')
})

test('readSheetView defaults gridlines and headings on and reads only the first view', () => {
  const bytes = encode(
    '<worksheet><sheetViews>' +
      '<sheetView zoomScale="x"/><sheetView showGridLines="0"/>' + // second view ignored; bad zoom dropped
      '</sheetViews><sheetData/></worksheet>',
  )
  const view = readSheetView(bytes)
  assert.equal(view.gridlines, true)
  assert.equal(view.headings, true)
  assert.equal(view.zoom, undefined)
  assert.equal(view.frozen, undefined)
  assert.equal(view.tabColor, undefined)
  assert.equal(view.autoFilter, undefined)
})

test('readHiddenRows and readHiddenColumns read the hidden flag', () => {
  const bytes = encode(
    '<worksheet><cols><col min="2" max="3" hidden="1"/><col min="5" max="5" width="8"/></cols>' +
      '<sheetData><row r="1" hidden="1"/><row r="2"/></sheetData></worksheet>',
  )
  const rows = readHiddenRows(bytes)
  assert.equal(rows.has(1), true)
  assert.equal(rows.has(2), false)
  assert.deepEqual(readHiddenColumns(bytes), [{ min: 2, max: 3 }])
})

test('readRowGroupLevels and readColumnGroupLevels read outlineLevel, skipping zero', () => {
  const bytes = encode(
    '<worksheet><cols><col min="2" max="3" outlineLevel="2"/><col min="5" max="5" outlineLevel="0"/></cols>' +
      '<sheetData><row r="1" outlineLevel="1"/><row r="2"/><row r="3" outlineLevel="0"/></sheetData></worksheet>',
  )
  const rows = readRowGroupLevels(bytes)
  assert.equal(rows.get(1), 1)
  assert.equal(rows.has(2), false)
  assert.equal(rows.has(3), false)
  assert.deepEqual(readColumnGroupLevels(bytes), [{ min: 2, max: 3, level: 2 }])
})

test('readDataValidations reads type, sqref, allowBlank and formulas, skipping the incomplete', () => {
  const bytes = encode(
    '<worksheet><sheetData/><dataValidations count="3">' +
      '<dataValidation type="list" allowBlank="1" sqref="A1:A5"><formula1>"a,b,c"</formula1></dataValidation>' +
      '<dataValidation type="whole" operator="between" allowBlank="0" sqref="B1">' +
      '<formula1>1</formula1><formula2>10</formula2></dataValidation>' +
      '<dataValidation type="list" sqref="C1"/>' + // no formula1
      '<dataValidation sqref="D1"><formula1>1</formula1></dataValidation>' + // no type, skipped
      '</dataValidations></worksheet>',
  )
  const specs = readDataValidations(bytes)
  assert.equal(specs.length, 3)
  assert.deepEqual(specs[0], {
    type: 'list',
    sqref: 'A1:A5',
    allowBlank: true,
    operator: undefined,
    formula1: '"a,b,c"',
    formula2: undefined,
  })
  assert.deepEqual(specs[1], {
    type: 'whole',
    sqref: 'B1',
    allowBlank: false,
    operator: 'between',
    formula1: '1',
    formula2: '10',
  })
  assert.equal(specs[2]?.sqref, 'C1') // kept even with an empty formula
})

test('readConditionalFormats reads colour scales, cellIs and data bars', () => {
  const bytes = encode(
    '<worksheet><sheetData/>' +
      '<conditionalFormatting sqref="A1:A9"><cfRule type="colorScale" priority="1"><colorScale>' +
      '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
      '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1"><cfRule type="cellIs" operator="between" dxfId="3" priority="2">' +
      '<formula>1</formula><formula>10</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="C1:C5"><cfRule type="dataBar" priority="3"><dataBar>' +
      '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>' +
      '</worksheet>',
  )
  const specs = readConditionalFormats(bytes)
  assert.deepEqual(specs[0], {
    kind: 'colorScale',
    sqref: 'A1:A9',
    colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
  })
  assert.deepEqual(specs[1], {
    kind: 'cellIs',
    sqref: 'B1',
    operator: 'between',
    formulas: ['1', '10'],
    dxfId: 3,
  })
  assert.deepEqual(specs[2], { kind: 'dataBar', sqref: 'C1:C5', color: 'FF638EC6' })
})

test('readConditionalFormats reads expression, duplicate and unique rules', () => {
  const bytes = encode(
    '<worksheet><sheetData/>' +
      '<conditionalFormatting sqref="A1:A9"><cfRule type="expression" dxfId="1" priority="1">' +
      '<formula>$A1&gt;0</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1:B9"><cfRule type="duplicateValues" dxfId="2" priority="2"/></conditionalFormatting>' +
      '<conditionalFormatting sqref="C1:C9"><cfRule type="uniqueValues" dxfId="3" priority="3"/></conditionalFormatting>' +
      '</worksheet>',
  )
  assert.deepEqual(readConditionalFormats(bytes), [
    { kind: 'expression', sqref: 'A1:A9', formula: '$A1>0', dxfId: 1 },
    { kind: 'duplicateValues', sqref: 'B1:B9', dxfId: 2 },
    { kind: 'uniqueValues', sqref: 'C1:C9', dxfId: 3 },
  ])
})

test('readConditionalFormats reads top10 rank, bottom and percent', () => {
  const bytes = encode(
    '<worksheet><sheetData/>' +
      '<conditionalFormatting sqref="A1:A9"><cfRule type="top10" dxfId="1" priority="1" rank="3"/></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1:B9"><cfRule type="top10" dxfId="2" priority="2" rank="10" bottom="1" percent="1"/></conditionalFormatting>' +
      '</worksheet>',
  )
  assert.deepEqual(readConditionalFormats(bytes), [
    { kind: 'top10', sqref: 'A1:A9', rank: 3, bottom: false, percent: false, dxfId: 1 },
    { kind: 'top10', sqref: 'B1:B9', rank: 10, bottom: true, percent: true, dxfId: 2 },
  ])
})
