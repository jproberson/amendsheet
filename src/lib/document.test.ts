import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { buildVmlDrawing } from './comments.js'
import { readContainer, writeContainer } from './container.js'
import { type Cell, createWorkbook, readWorkbook } from './document.js'
import { XlsxError } from './errors.js'
import type { CellInput } from './patch.js'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array | undefined) =>
  new TextDecoder().decode(bytes ?? new Uint8Array())

const ROOT_RELS = `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
const WORKBOOK_RELS = `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`

const STYLES = `<styleSheet><numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>`

function build(
  sheetBody: string,
  options: {
    date1904?: boolean
    extra?: Record<string, string>
    sheetPr?: string
    views?: string
    after?: string
  } = {},
) {
  const parts: Record<string, Uint8Array> = {
    '_rels/.rels': encode(ROOT_RELS),
    'xl/workbook.xml': encode(
      `<workbook><workbookPr${options.date1904 ? ' date1904="1"' : ''}/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': encode(WORKBOOK_RELS),
    'xl/styles.xml': encode(STYLES),
    'xl/worksheets/sheet1.xml': encode(
      `<worksheet>${options.sheetPr ?? ''}${options.views ?? ''}<sheetData>${sheetBody}</sheetData>${options.after ?? ''}</worksheet>`,
    ),
    'xl/charts/chart1.xml': encode('<chart/>'),
  }
  for (const [path, content] of Object.entries(options.extra ?? {})) parts[path] = encode(content)
  return writeContainer({ parts: new Map(Object.entries(parts)) })
}

test('createWorkbook makes an empty one-sheet workbook that fills and writes back', () => {
  const workbook = createWorkbook()

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Sheet1'],
  )
  assert.deepEqual([...(workbook.sheets[0]?.cells() ?? [])], [])

  workbook.sheets[0]?.set('A1', 'hi')
  workbook.sheets[0]?.set('B2', 42, { font: { bold: true } })

  const back = readWorkbook(workbook.toBytes())
  const cell = (reference: string) =>
    [...(back.sheets[0]?.cells() ?? [])].find((c) => c.reference === reference)
  assert.deepEqual(cell('A1')?.value, { kind: 'text', value: 'hi' })
  assert.deepEqual(cell('B2')?.value, { kind: 'number', value: 42 })
})

test('createWorkbook names its sheet and refuses a bad name', () => {
  assert.equal(createWorkbook().sheets[0]?.name, 'Sheet1')
  const named = createWorkbook('Budget')
  assert.equal(named.sheets[0]?.name, 'Budget')
  assert.equal(readWorkbook(named.toBytes()).sheets[0]?.name, 'Budget')
  assert.throws(
    () => createWorkbook('bad/name'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('addSheet adds editable sheets a created workbook writes back', () => {
  const workbook = createWorkbook()
  workbook.sheets[0]?.set('A1', 'first')
  const data = workbook.addSheet('Data')
  data.set('A1', 'second')
  data.set('B2', 99, { font: { bold: true } })

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Sheet1', 'Data'],
  )
  assert.equal(workbook.sheet('Data'), data)

  const back = readWorkbook(workbook.toBytes())
  assert.deepEqual(
    back.sheets.map((sheet) => sheet.name),
    ['Sheet1', 'Data'],
  )
  const cell = (sheet: number, reference: string) =>
    [...(back.sheets[sheet]?.cells() ?? [])].find((c) => c.reference === reference)
  assert.deepEqual(cell(0, 'A1')?.value, { kind: 'text', value: 'first' })
  assert.deepEqual(cell(1, 'A1')?.value, { kind: 'text', value: 'second' })
  assert.deepEqual(cell(1, 'B2')?.value, { kind: 'number', value: 99 })
})

test('addSheet adds a sheet to a workbook that was read, leaving other parts alone', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.addSheet('Extra').set('A1', 'new')

  const out = workbook.toBytes()
  const parts = readContainer(out).parts
  assert.equal(decode(parts.get('xl/charts/chart1.xml')), '<chart/>')

  const back = readWorkbook(out)
  assert.deepEqual(
    back.sheets.map((sheet) => sheet.name),
    ['Data', 'Extra'],
  )
  const extra = [...(back.sheets[1]?.cells() ?? [])].find((c) => c.reference === 'A1')
  assert.deepEqual(extra?.value, { kind: 'text', value: 'new' })
})

test('addSheet refuses a name Excel will not take', () => {
  const workbook = createWorkbook()
  const refuses = (name: string) =>
    assert.throws(
      () => workbook.addSheet(name),
      (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
    )
  refuses('Sheet1') // a name a sheet already uses
  refuses('sheet1') // the same name in another case
  refuses('') // empty
  refuses('a'.repeat(32)) // longer than 31 characters
  refuses('bad/name') // a character Excel forbids
})

test('rename changes a sheet name on both a read and an added sheet', () => {
  const workbook = createWorkbook('Old')
  workbook.sheets[0]?.set('A1', 'keep')
  workbook.sheets[0]?.rename('Renamed')
  const added = workbook.addSheet('Temp')
  added.set('A1', 'x')
  added.rename('Final')

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Renamed', 'Final'],
  )
  assert.equal(workbook.sheet('Renamed'), workbook.sheets[0])

  const back = readWorkbook(workbook.toBytes())
  assert.deepEqual(
    back.sheets.map((sheet) => sheet.name),
    ['Renamed', 'Final'],
  )
  assert.deepEqual(
    [...(back.sheets[0]?.cells() ?? [])].map((c) => c.reference),
    ['A1'],
  )
})

test('rename refuses a duplicate but allows a sheet its own name', () => {
  const workbook = createWorkbook('A')
  workbook.addSheet('B')
  assert.throws(
    () => workbook.sheets[0]?.rename('B'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  workbook.sheets[0]?.rename('A')
  assert.equal(workbook.sheets[0]?.name, 'A')
})

test('remove drops a sheet and refuses the last one', () => {
  const workbook = createWorkbook('One')
  workbook.sheets[0]?.set('A1', 'keep')
  workbook.addSheet('Two').set('A1', 'gone')
  workbook.addSheet('Three').set('A1', 'stay')

  workbook.sheet('Two')?.remove()
  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['One', 'Three'],
  )

  const back = readWorkbook(workbook.toBytes())
  assert.deepEqual(
    back.sheets.map((sheet) => sheet.name),
    ['One', 'Three'],
  )
  assert.deepEqual(
    [...(back.sheet('Three')?.cells() ?? [])].map((c) => c.reference),
    ['A1'],
  )

  const solo = createWorkbook()
  assert.throws(
    () => solo.sheets[0]?.remove(),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('remove drops a sheet the file was read with, and its part', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.addSheet('Extra')
  workbook.sheet('Data')?.remove()

  const out = workbook.toBytes()
  assert.ok(![...readContainer(out).parts.keys()].includes('xl/worksheets/sheet1.xml'))
  assert.deepEqual(
    readWorkbook(out).sheets.map((sheet) => sheet.name),
    ['Extra'],
  )
})

test('autoFilter writes the sheet filter, canonicalising the range', () => {
  const workbook = createWorkbook()
  workbook.sheets[0]?.autoFilter('a1:c1')

  const out = workbook.toBytes()
  assert.match(
    decode(readContainer(out).parts.get('xl/worksheets/sheet1.xml')),
    /<autoFilter ref="A1:C1"\/>/,
  )
  readWorkbook(out) // the result stays a readable workbook
})

test('freeze writes a frozen pane, canonicalising the cell', () => {
  const workbook = createWorkbook()
  workbook.sheets[0]?.freeze('b2')

  const out = workbook.toBytes()
  assert.match(
    decode(readContainer(out).parts.get('xl/worksheets/sheet1.xml')),
    /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/,
  )
  readWorkbook(out)
})

test('hideRow and hideColumn write the hidden flag, refusing a bad reference', () => {
  const workbook = createWorkbook()
  workbook.sheets[0]?.hideRow(2)
  workbook.sheets[0]?.hideColumn('B')

  const out = workbook.toBytes()
  const xml = decode(readContainer(out).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(xml, /<row r="2" hidden="1">/)
  assert.match(xml, /<col min="2" max="2" hidden="1"\/>/)
  readWorkbook(out)

  assert.throws(
    () => workbook.sheets[0]?.hideRow(0),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => workbook.sheets[0]?.hideColumn('not a column'),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('defineName creates and reads back a global name, replacing a redefine', () => {
  const workbook = createWorkbook()
  workbook.defineName('Tax', 'Sheet1!$B$1')
  assert.deepEqual([...workbook.definedNames], [['Tax', 'Sheet1!$B$1']])

  const back = readWorkbook(workbook.toBytes())
  assert.deepEqual([...back.definedNames], [['Tax', 'Sheet1!$B$1']])

  back.defineName('Tax', 'Sheet1!$C$1')
  assert.deepEqual([...readWorkbook(back.toBytes()).definedNames], [['Tax', 'Sheet1!$C$1']])
})

test('exposes sheets by name', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Data'],
  )
})

test('a malformed cell reference in the file reads as a file fault, not a caller fault', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A"><v>1</v></c></row>'))

  assert.throws(
    () => [...(workbook.sheets[0]?.cells() ?? [])],
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'invalid-content' &&
      error.reference === 'A' &&
      error.sheet === 'Data' &&
      error.part === 'xl/worksheets/sheet1.xml',
  )
})

test('reads a number as a number', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>42</v></c></row>'))
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'number', value: 42 })
})

test('reads a number with a date format as a date', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'))
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind, 'date')
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getFullYear(), 2024)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getMonth(), 0)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getDate(), 1)
})

test('keeps the serial alongside a date', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'))
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind === 'date' && cell.value.serial, 45292)
})

test('reads dates against the workbook epoch', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="1"><v>0</v></c></row>', { date1904: true }),
  )
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind === 'date' && cell.value.value.getFullYear(), 1904)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getDate(), 1)
})

test('leaves a date formatted cell that holds text alone', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>not a date</t></is></c></row>'),
  )
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'text', value: 'not a date' })
})

test('reads shared strings through the workbook', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', {
      extra: {
        'xl/sharedStrings.xml': '<sst><si><t>shared</t></si></sst>',
      },
    }),
  )
  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'text', value: 'shared' })
})

test('reads a workbook that has no styles part', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
      [
        'xl/worksheets/sheet1.xml',
        encode(
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        ),
      ],
    ]),
  })

  const [cell] = [...(readWorkbook(bytes).sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'number', value: 1 })
})

test('reports no cells for a sheet whose part is missing', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
    ]),
  })

  assert.deepEqual([...(readWorkbook(bytes).sheets[0]?.cells() ?? [])], [])
})

test('writes back every part it read, including ones it does not understand', () => {
  const original = build('<row r="1"><c r="A1"><v>1</v></c></row>')

  const rewritten = readContainer(readWorkbook(original).toBytes())

  assert.equal(rewritten.parts.has('xl/charts/chart1.xml'), true)
  assert.deepEqual(
    [...rewritten.parts.keys()].sort(),
    [...readContainer(original).parts.keys()].sort(),
  )
})

test('round trips every fixtures file without losing a part', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  const damaged: string[] = []

  for (const file of files) {
    const original = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const before = readContainer(original)
    const after = readContainer(readWorkbook(original).toBytes())

    for (const [path, bytes] of before.parts) {
      const other = after.parts.get(path)
      if (other === undefined) damaged.push(`${file}: lost ${path}`)
      else if (Buffer.compare(Buffer.from(bytes), Buffer.from(other)) !== 0) {
        damaged.push(`${file}: changed ${path}`)
      }
    }
  }

  assert.deepEqual(damaged, [])
})

test('reads cells from every fixtures file', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  let cells = 0
  let dates = 0

  for (const file of files) {
    const workbook = readWorkbook(new Uint8Array(await readFile(`fixtures/real/${file}`)))
    for (const sheet of workbook.sheets) {
      for (const cell of sheet.cells()) {
        cells++
        if (cell.value.kind === 'date') dates++
      }
    }
  }

  assert.ok(cells > 25000, `expected many cells, got ${cells}`)
  assert.ok(dates > 0, `expected some dates, got ${dates}`)
})

test('reads back a value that was set', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 42)

  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'number', value: 42 })
})

test('writes a value that was set', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'changed')

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'text', value: 'changed' })
})

test('leaves parts it does not touch alone when a value is set', () => {
  const original = build('<row r="1"><c r="A1"><v>1</v></c></row>')
  const workbook = readWorkbook(original)
  workbook.sheets[0]?.set('A1', 2)

  const before = readContainer(original)
  const after = readContainer(workbook.toBytes())

  for (const [path, bytes] of before.parts) {
    if (path === 'xl/worksheets/sheet1.xml') continue
    assert.deepEqual(after.parts.get(path), bytes, `${path} changed`)
  }
})

test('rejects a reference that is not a cell', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(() => workbook.sheets[0]?.set('nonsense', 1), /not a cell reference/)
})

test('writes a date into a cell that already has a date format', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', new Date(2024, 0, 1))

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind, 'date')
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getFullYear(), 2024)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getMonth(), 0)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getDate(), 1)
})

const withStrings = (sheetBody: string, sst: string) =>
  build(sheetBody, { extra: { 'xl/sharedStrings.xml': sst } })

test('set applies a font, adding it to the styles and pointing the cell at it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'bold', { font: { bold: true, color: 'FF0000' } })

  const parts = readContainer(workbook.toBytes()).parts
  const styles = new TextDecoder().decode(parts.get('xl/styles.xml') ?? new Uint8Array())
  const sheet = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml') ?? new Uint8Array())

  assert.match(styles, /<font><b\/><color rgb="FFFF0000"\/><\/font>/)
  assert.match(sheet, /<c r="A1"[^>]* s="\d+"/)
})

test('set composes a number format and a font into one cell format', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 5, { numberFormat: '0.00', font: { italic: true } })

  const styles = new TextDecoder().decode(
    readContainer(workbook.toBytes()).parts.get('xl/styles.xml') ?? new Uint8Array(),
  )

  assert.match(styles, /<font><i\/><\/font>/)
  const xf = [...styles.matchAll(/<xf [^>]*\/>/g)].map((match) => match[0]).at(-1) ?? ''
  assert.match(xf, /applyNumberFormat="1"/)
  assert.match(xf, /applyFont="1"/)
})

test('set refuses a font colour that is not hex, before recording anything', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.throws(() => sheet?.set('A1', 'x', { font: { color: 'nope' } }), /hex/)
  // The refused edit left no trace: A1 still reads its original number.
  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
})

const locates = (sheet: string, reference: string) => (error: unknown) =>
  error instanceof XlsxError &&
  error.code === 'unwritable-value' &&
  error.sheet === sheet &&
  error.reference === reference

test('a refused font colour names the sheet and cell it was for', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(() => sheet?.set('A1', 'x', { font: { color: 'nope' } }), locates('Data', 'A1'))
})

test('a refused text rotation names the sheet and cell it was for', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () => sheet?.set('A1', 'x', { alignment: { textRotation: 999 } }),
    locates('Data', 'A1'),
  )
})

const THEME1 = `<theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><themeElements><clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></clrScheme></themeElements></theme>`

test('resolveColor resolves a theme colour against the workbook theme part', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { extra: { 'xl/theme/theme1.xml': THEME1 } }),
  )
  assert.equal(workbook.resolveColor({ theme: 4 }), 'FF4472C4')
  assert.equal(workbook.resolveColor({ theme: 4, tint: -0.499985 }), 'FF203764')
})

test('resolveColor without a theme part still resolves hex and indexed colours', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  assert.equal(workbook.resolveColor('4472C4'), 'FF4472C4')
  assert.equal(workbook.resolveColor({ indexed: 10 }), 'FFFF0000')
  // No theme, so a theme reference has nothing to resolve against.
  assert.equal(workbook.resolveColor({ theme: 4 }), undefined)
})

const lastXf = (styles: string) =>
  [...styles.matchAll(/<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g)]
    .map((match) => match[0])
    .at(-1) ?? ''

const sheetXml = (workbook: ReturnType<typeof readWorkbook>) =>
  decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))

test('tabColor writes a tabColor into a fresh sheetPr, normalized to 8-digit ARGB', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.tabColor('FF0000')

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetPr><tabColor rgb="FFFF0000"\/><\/sheetPr>/)
  // sheetPr must be the first child of the worksheet.
  assert.match(xml, /<worksheet><sheetPr>/)
})

test('tabColor replaces a tabColor the sheet already has, keeping other sheetPr children', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      sheetPr: '<sheetPr codeName="S1"><tabColor rgb="FF00FF00"/></sheetPr>',
    }),
  )
  workbook.sheets[0]?.tabColor('#0000FF')

  const xml = sheetXml(workbook)
  assert.match(xml, /<tabColor rgb="FF0000FF"\/>/)
  assert.doesNotMatch(xml, /FF00FF00/)
  assert.match(xml, /codeName="S1"/)
})

test('tabColor opens a self-closing sheetPr to hold the colour', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { sheetPr: '<sheetPr codeName="S1"/>' }),
  )
  workbook.sheets[0]?.tabColor('FF0000')

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetPr codeName="S1"><tabColor rgb="FFFF0000"\/><\/sheetPr>/)
})

test('tabColor slots in as the first child of a sheetPr that has other children', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      sheetPr: '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>',
    }),
  )
  workbook.sheets[0]?.tabColor('FF0000')

  const xml = sheetXml(workbook)
  // tabColor must precede outlinePr per the sheetPr child order.
  assert.match(xml, /<sheetPr><tabColor rgb="FFFF0000"\/><outlinePr summaryBelow="0"\/><\/sheetPr>/)
})

test('tabColor refuses a colour that is not hex, naming the sheet', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () => sheet?.tabColor('reddish'),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data',
  )
})

test('showGridlines, showHeadings and zoom open a fresh sheetView with the flags', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.showGridlines(false)
  sheet?.showHeadings(false)
  sheet?.zoom(85)

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<sheetViews><sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="0" zoomScale="85"\/><\/sheetViews>/,
  )
})

test('a view flag and a freeze fold into one sheetView', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.showGridlines(false)
  workbook.sheets[0]?.freeze('B2')

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetView workbookViewId="0" showGridLines="0">/)
  assert.match(
    xml,
    /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/,
  )
})

test('zoom adds an attribute to a self-closing sheetView the sheet already has', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views: '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>',
    }),
  )
  workbook.sheets[0]?.zoom(120)

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetView[^>]*zoomScale="120"[^>]*\/>/)
  assert.match(xml, /tabSelected="1"/)
})

test('a view flag keeps the other children of an existing open sheetView', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views:
        '<sheetViews><sheetView workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews>',
    }),
  )
  workbook.sheets[0]?.showHeadings(false)

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetView[^>]*showRowColHeaders="0"[^>]*>/)
  assert.match(xml, /<selection activeCell="C3" sqref="C3"\/>/)
})

test('zoom on a sheetView that already holds a freeze pane keeps the pane', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views:
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews>',
    }),
  )
  workbook.sheets[0]?.zoom(75)

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetView[^>]*zoomScale="75"[^>]*>/)
  assert.match(xml, /<pane ySplit="1" topLeftCell="A2" state="frozen"\/>/)
})

test('a freeze and a zoom compose on an existing open sheetView, pane before its selection', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views:
        '<sheetViews><sheetView workbookViewId="0"><selection sqref="A1"/></sheetView></sheetViews>',
    }),
  )
  workbook.sheets[0]?.zoom(90)
  workbook.sheets[0]?.freeze('A2')

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetView[^>]*zoomScale="90"[^>]*>/)
  // pane is the first child, ahead of the selection the sheet already had.
  assert.match(
    xml,
    /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/><selection sqref="A1"\/>/,
  )
})

test('a freeze and a view flag open a self-closing sheetView together', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views: '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
    }),
  )
  workbook.sheets[0]?.showGridlines(false)
  workbook.sheets[0]?.freeze('A2')

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<sheetView[^>]*showGridLines="0"[^>]*><pane [^>]*state="frozen"\/><\/sheetView>/,
  )
})

test('zoom refuses a value outside the allowed range, naming the sheet', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () => sheet?.zoom(5),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data',
  )
})

test('groupRows sets an outline level on each row in the range, creating absent ones', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.groupRows(2, 4)

  const xml = sheetXml(workbook)
  assert.match(xml, /<row r="2"[^>]*outlineLevel="1"/)
  assert.match(xml, /<row r="3"[^>]*outlineLevel="1"/)
  assert.match(xml, /<row r="4"[^>]*outlineLevel="1"/)
})

test('groupRows adds the level to a row the sheet already has', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'),
  )
  workbook.sheets[0]?.groupRows(2, 2)

  const xml = sheetXml(workbook)
  assert.match(xml, /<row [^>]*outlineLevel="1"[^>]*><c r="A2"><v>2<\/v><\/c><\/row>/)
})

test('groupColumns sets an outline level on the columns in the range', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.groupColumns('B', 'D')

  const xml = sheetXml(workbook)
  assert.match(xml, /<col min="2" max="2"[^>]*outlineLevel="1"\/>/)
  assert.match(xml, /<col min="4" max="4"[^>]*outlineLevel="1"\/>/)
})

test('a nested level updates an existing sheetFormatPr outline hint', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views: '<sheetFormatPr defaultRowHeight="15"/>',
    }),
  )
  workbook.sheets[0]?.groupRows(2, 3, 2)

  const xml = sheetXml(workbook)
  assert.match(xml, /<sheetFormatPr[^>]*outlineLevelRow="2"[^>]*\/>/)
  assert.match(xml, /<row r="2"[^>]*outlineLevel="2"/)
})

test('groupRows refuses a backwards range and a level out of bounds, naming the sheet', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  const refused = (error: unknown) =>
    error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data'
  assert.throws(() => sheet?.groupRows(4, 2), refused)
  assert.throws(() => sheet?.groupRows(2, 4, 9), refused)
})

test('groupColumns refuses a backwards range, a fractional and a too-low level', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  const refused = (error: unknown) =>
    error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data'
  assert.throws(() => sheet?.groupColumns('D', 'B'), refused)
  assert.throws(() => sheet?.groupColumns('B', 'D', 1.5), refused)
  assert.throws(() => sheet?.groupColumns('B', 'D', 0), refused)
})

test('the outline hint keeps a deeper level the sheet declared and gains a column level', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      views: '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="3"/>',
    }),
  )
  workbook.sheets[0]?.groupRows(2, 2, 1)
  workbook.sheets[0]?.groupColumns('B', 'C', 2)

  const xml = sheetXml(workbook)
  // The row level the file declared (3) outranks the one grouped now (1).
  assert.match(xml, /outlineLevelRow="3"/)
  assert.match(xml, /outlineLevelCol="2"/)
})

test('validate writes an inline list dropdown as a dataValidation before the worksheet close', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.validate('B2:B10', { list: ['Yes', 'No', 'Maybe'] })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2:B10"><formula1>"Yes,No,Maybe"<\/formula1><\/dataValidation><\/dataValidations><\/worksheet>/,
  )
})

test('a dataValidation lands ahead of a later sibling like pageMargins', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { after: '<pageMargins left="0.7"/>' }),
  )
  workbook.sheets[0]?.validate('A1', { list: ['x', 'y'] })

  const xml = sheetXml(workbook)
  assert.match(xml, /<\/dataValidations><pageMargins left="0.7"\/>/)
})

test('a second validate joins the dataValidations the sheet already has', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after:
        '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"a,b"</formula1></dataValidation></dataValidations>',
    }),
  )
  workbook.sheets[0]?.validate('B2:B5', { list: ['c', 'd'] })

  const xml = sheetXml(workbook)
  assert.match(xml, /<dataValidations count="2">/)
  assert.match(xml, /sqref="A1"/)
  assert.match(xml, /sqref="B2:B5"/)
})

test('validate writes a whole-number between constraint with two formulas', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.validate('C2:C10', { whole: { between: [1, 100] } })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C2:C10"><formula1>1<\/formula1><formula2>100<\/formula2><\/dataValidation>/,
  )
})

test('validate writes a decimal greaterThan constraint with one formula', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.validate('D1', { decimal: { greaterThan: 3.5 } })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<dataValidation type="decimal" operator="greaterThan"[^>]*sqref="D1"><formula1>3.5<\/formula1><\/dataValidation>/,
  )
  assert.doesNotMatch(xml, /formula2/)
})

test('validate covers each numeric operator', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.validate('A1', { whole: { notBetween: [1, 2] } })
  sheet?.validate('A2', { whole: { equal: 5 } })
  sheet?.validate('A3', { whole: { notEqual: 5 } })
  sheet?.validate('A4', { whole: { lessThan: 5 } })
  sheet?.validate('A5', { whole: { greaterThanOrEqual: 5 } })
  sheet?.validate('A6', { whole: { lessThanOrEqual: 5 } })

  const xml = sheetXml(workbook)
  for (const operator of [
    'notBetween',
    'equal',
    'notEqual',
    'lessThan',
    'greaterThanOrEqual',
    'lessThanOrEqual',
  ]) {
    assert.match(xml, new RegExp(`operator="${operator}"`))
  }
})

test('validate refuses a non-finite numeric bound, naming the sheet', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () => sheet?.validate('A1', { whole: { equal: Number.POSITIVE_INFINITY } }),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data',
  )
})

test('validate opens a self-closing dataValidations the sheet had', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { after: '<dataValidations count="0"/>' }),
  )
  workbook.sheets[0]?.validate('B2', { list: ['x'] })

  const xml = sheetXml(workbook)
  assert.match(xml, /<dataValidations count="1"><dataValidation [^>]*sqref="B2">/)
})

test('validate allowBlank false and escapes list values', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.validate('C1', { list: ['a<b', 'c&d'], allowBlank: false })

  const xml = sheetXml(workbook)
  assert.match(xml, /allowBlank="0"/)
  assert.match(xml, /<formula1>"a&lt;b,c&amp;d"<\/formula1>/)
})

test('validate refuses an empty list, a value with a comma, and a bad range', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  const refused = (error: unknown) =>
    error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data'
  assert.throws(() => sheet?.validate('A1', { list: [] }), refused)
  assert.throws(() => sheet?.validate('A1', { list: ['a,b'] }), refused)
  assert.throws(
    () => sheet?.validate('not a range', { list: ['a'] }),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('conditionalFormat writes a two-colour scale', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.conditionalFormat('A1:A20', { colorScale: { min: 'FFFFFF', max: 'FF0000' } })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<conditionalFormatting sqref="A1:A20"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"\/><cfvo type="max"\/><color rgb="FFFFFFFF"\/><color rgb="FFFF0000"\/><\/colorScale><\/cfRule><\/conditionalFormatting>/,
  )
})

test('conditionalFormat writes a three-colour scale with a midpoint', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.conditionalFormat('B1:B5', {
    colorScale: { min: 'F8696B', mid: 'FFEB84', max: '63BE7B' },
  })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<colorScale><cfvo type="min"\/><cfvo type="percentile" val="50"\/><cfvo type="max"\/><color rgb="FFF8696B"\/><color rgb="FFFFEB84"\/><color rgb="FF63BE7B"\/><\/colorScale>/,
  )
})

test('conditionalFormat writes a data bar with min and max stops', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.conditionalFormat('A1:A20', { dataBar: { color: '638EC6' } })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF638EC6"\/><\/dataBar><\/cfRule>/,
  )
})

test('a conditional format lands before dataValidations', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { after: '<dataValidations count="0"/>' }),
  )
  workbook.sheets[0]?.conditionalFormat('A1', { colorScale: { min: 'FFFFFF', max: '000000' } })

  const xml = sheetXml(workbook)
  assert.match(xml, /<\/conditionalFormatting><dataValidations/)
})

test('a conditional format priority rises above the highest the sheet already uses', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after:
        '<conditionalFormatting sqref="Z1"><cfRule type="expression" priority="5"><formula>TRUE</formula></cfRule></conditionalFormatting>',
    }),
  )
  workbook.sheets[0]?.conditionalFormat('A1', { colorScale: { min: 'FFFFFF', max: '000000' } })

  const xml = sheetXml(workbook)
  assert.match(xml, /sqref="A1"><cfRule type="colorScale" priority="6"/)
})

test('conditionalFormat writes a cellIs rule and a dxf holding its fill', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.conditionalFormat('A1:A20', {
    cellIs: { when: { greaterThan: 100 }, fill: 'FFFF00' },
  })

  const parts = readContainer(workbook.toBytes()).parts
  assert.match(
    decode(parts.get('xl/worksheets/sheet1.xml')),
    /<cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>100<\/formula><\/cfRule>/,
  )
  assert.match(
    decode(parts.get('xl/styles.xml')),
    /<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFFF00"\/><\/patternFill><\/fill><\/dxf><\/dxfs><\/styleSheet>/,
  )
})

test('a cellIs between rule writes two formulas', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.conditionalFormat('B1', {
    cellIs: { when: { between: [1, 10] }, fill: 'FF0000' },
  })

  const xml = sheetXml(workbook)
  assert.match(
    xml,
    /<cfRule type="cellIs" operator="between" dxfId="0" priority="1"><formula>1<\/formula><formula>10<\/formula><\/cfRule>/,
  )
})

test('a cellIs rule appends to a dxfs table the styles already have', () => {
  const styles =
    '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs>' +
    '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )
  workbook.sheets[0]?.conditionalFormat('A1', { cellIs: { when: { equal: 0 }, fill: '00FF00' } })

  const parts = readContainer(workbook.toBytes()).parts
  assert.match(decode(parts.get('xl/styles.xml')), /<dxfs count="2">/)
  assert.match(decode(parts.get('xl/worksheets/sheet1.xml')), /dxfId="1"/)
})

test('a cellIs rule opens a self-closing dxfs and lands before tableStyles', () => {
  const styles =
    '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs><dxfs count="0"/>' +
    '<tableStyles count="0"/></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )
  workbook.sheets[0]?.conditionalFormat('A1', { cellIs: { when: { equal: 1 }, fill: 'FF0000' } })

  const out = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(out, /<dxfs count="1"><dxf>.*<\/dxf><\/dxfs><tableStyles count="0"\/>/)
})

test('a cellIs rule inserts a fresh dxfs before an existing tableStyles', () => {
  const styles =
    '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs>' +
    '<tableStyles count="0"/></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )
  workbook.sheets[0]?.conditionalFormat('A1', { cellIs: { when: { equal: 1 }, fill: 'FF0000' } })

  const out = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(out, /<\/cellXfs><dxfs count="1"><dxf>.*<\/dxf><\/dxfs><tableStyles/)
})

test('a cellIs rule needs a style table to hold its fill', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode(
          '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        ),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
      ['xl/worksheets/sheet1.xml', encode('<worksheet><sheetData/></worksheet>')],
    ]),
  })
  const sheet = readWorkbook(bytes).sheet('Data')
  assert.throws(
    () => sheet?.conditionalFormat('A1', { cellIs: { when: { equal: 1 }, fill: 'FF0000' } }),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
})

test('conditionalFormat refuses a cellIs bound that is not finite', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () =>
      sheet?.conditionalFormat('A1', { cellIs: { when: { equal: Number.NaN }, fill: 'FF0000' } }),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data',
  )
})

test('conditionalFormat refuses a bad colour and a bad range', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheet('Data')
  assert.throws(
    () => sheet?.conditionalFormat('A1', { colorScale: { min: 'nope', max: '000000' } }),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'unwritable-value' && error.sheet === 'Data',
  )
  assert.throws(
    () => sheet?.conditionalFormat('bad range', { colorScale: { min: 'FFFFFF', max: '000000' } }),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('cell reports a gradient fill, which set itself does not write', () => {
  const styles =
    '<styleSheet><fills count="3">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><gradientFill degree="90">' +
    '<stop position="0"><color rgb="FFFF0000"/></stop>' +
    '<stop position="1"><color rgb="FF0000FF"/></stop>' +
    '</gradientFill></fill></fills>' +
    '<cellXfs count="1"><xf fillId="2"/></cellXfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="0"><v>1</v></c></row>', {
      extra: { 'xl/styles.xml': styles },
    }),
  )

  assert.deepEqual(workbook.sheets[0]?.cell('A1')?.fill, {
    type: 'gradient',
    degree: 90,
    stops: [
      { position: 0, color: 'FFFF0000' },
      { position: 1, color: 'FF0000FF' },
    ],
  })
})

test('a gradient with no degree is reported without one', () => {
  const styles =
    '<styleSheet><fills count="1">' +
    '<fill><gradientFill type="path">' +
    '<stop position="0"><color rgb="FF000000"/></stop>' +
    '<stop position="1"><color rgb="FFFFFFFF"/></stop>' +
    '</gradientFill></fill></fills>' +
    '<cellXfs count="1"><xf fillId="0"/></cellXfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="0"><v>1</v></c></row>', {
      extra: { 'xl/styles.xml': styles },
    }),
  )

  const fill = workbook.sheets[0]?.cell('A1')?.fill
  assert.deepEqual(fill, {
    type: 'gradient',
    stops: [
      { position: 0, color: 'FF000000' },
      { position: 1, color: 'FFFFFFFF' },
    ],
  })
})

const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const commentsRels = (target: string) =>
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId9" Type="${COMMENTS_REL}" Target="${target}"/></Relationships>`

test('cell reports a comment wired through the sheet relationships', () => {
  const comments =
    '<comments><authors><author>Ada</author></authors><commentList>' +
    '<comment ref="B2" authorId="0"><text><r><t>Look </t></r><r><t>here</t></r></text></comment>' +
    '</commentList></comments>'
  const workbook = readWorkbook(
    build('<row r="2"><c r="B2"><v>5</v></c></row>', {
      extra: {
        'xl/comments1.xml': comments,
        'xl/worksheets/_rels/sheet1.xml.rels': commentsRels('../comments1.xml'),
      },
    }),
  )
  assert.equal(workbook.sheets[0]?.cell('B2')?.comment, 'Look here')
})

test('a comment on a cell the sheet never gave a value is still read', () => {
  const comments =
    '<comments><authors><author/></authors><commentList>' +
    '<comment ref="D4" authorId="0"><text><r><t>note on an empty cell</t></r></text></comment>' +
    '</commentList></comments>'
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/comments1.xml': comments,
        'xl/worksheets/_rels/sheet1.xml.rels': commentsRels('../comments1.xml'),
      },
    }),
  ).sheet('Data')

  const cell = sheet?.cell('D4')
  assert.equal(cell?.comment, 'note on an empty cell')
  assert.deepEqual(cell?.value, { kind: 'empty' })
  assert.ok([...(sheet?.cells() ?? [])].some((c) => c.reference === 'D4'))
})

test('mergedRanges reports the sheet merges and any added this session', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after: '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>',
    }),
  ).sheets[0]
  assert.deepEqual(sheet?.mergedRanges, ['A1:B2'])
  sheet?.merge('d4:e5')
  assert.deepEqual(sheet?.mergedRanges, ['A1:B2', 'D4:E5'])
})

test('mergedRanges is empty for a sheet with no merges', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheets[0]
  assert.deepEqual(sheet?.mergedRanges, [])
})

test('validations round-trip a list, whole and decimal rule', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.validate('A1:A5', { list: ['red', 'green'] })
  sheet?.validate('B1', { whole: { between: [1, 10] } })
  sheet?.validate('C1', { decimal: { greaterThan: 0 }, allowBlank: false })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.validations, [
    { range: 'A1:A5', rule: { allowBlank: true, list: ['red', 'green'] } },
    { range: 'B1', rule: { allowBlank: true, whole: { between: [1, 10] } } },
    { range: 'C1', rule: { allowBlank: false, decimal: { greaterThan: 0 } } },
  ])
})

test('validations round-trip every numeric operator', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  const wholes = [
    { equal: 1 },
    { notEqual: 2 },
    { greaterThan: 3 },
    { lessThan: 4 },
    { greaterThanOrEqual: 5 },
    { lessThanOrEqual: 6 },
    { between: [7, 8] },
    { notBetween: [9, 10] },
  ] as const
  const refs = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8']
  wholes.forEach((whole, i) => {
    sheet?.validate(refs[i] ?? 'A1', { whole })
  })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(
    back?.validations.map((v) => v.rule),
    wholes.map((whole) => ({ allowBlank: true, whole })),
  )
})

test('validations reflect a pending validate and skip a kind not modelled', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after:
        '<dataValidations count="4">' +
        '<dataValidation type="time" operator="lessThan" sqref="A1"><formula1>0.5</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="wonky" sqref="C1"><formula1>1</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="equal" sqref="D1"><formula1>abc</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="between" sqref="E1"><formula1>1</formula1><formula2>x</formula2></dataValidation>' +
        '</dataValidations>',
    }),
  ).sheets[0]
  // A time, an unknown operator and non-numeric bounds are all left out.
  assert.deepEqual(sheet?.validations, [])
  sheet?.validate('F1', { whole: { equal: 7 } })
  assert.deepEqual(sheet?.validations, [
    { range: 'F1', rule: { allowBlank: true, whole: { equal: 7 } } },
  ])
})

test('validations round-trip textLength and custom rules', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.validate('A1', { textLength: { lessThanOrEqual: 10 } })
  sheet?.validate('B1', { custom: 'LEN(B1)>3', allowBlank: false })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.validations, [
    { range: 'A1', rule: { allowBlank: true, textLength: { lessThanOrEqual: 10 } } },
    { range: 'B1', rule: { allowBlank: false, custom: 'LEN(B1)>3' } }, // formula's > survives escaping
  ])
})

test('validations round-trip a range-backed list and a date rule', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.validate('A1', { listRange: 'Sheet1!$A$1:$A$3' })
  sheet?.validate('B1', {
    date: { between: [new Date(Date.UTC(2020, 0, 1)), new Date(Date.UTC(2020, 11, 31))] },
  })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.validations, [
    { range: 'A1', rule: { allowBlank: true, listRange: 'Sheet1!$A$1:$A$3' } },
    {
      range: 'B1',
      rule: {
        allowBlank: true,
        date: { between: [new Date(Date.UTC(2020, 0, 1)), new Date(Date.UTC(2020, 11, 31))] },
      },
    },
  ])
})

test('validations round-trip every date operator', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  const day = (n: number) => new Date(Date.UTC(2020, 0, n))
  const dates = [
    { equal: day(1) },
    { notEqual: day(2) },
    { greaterThan: day(3) },
    { lessThan: day(4) },
    { greaterThanOrEqual: day(5) },
    { lessThanOrEqual: day(6) },
    { between: [day(7), day(8)] as [Date, Date] },
    { notBetween: [day(9), day(10)] as [Date, Date] },
  ]
  const refs = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8']
  dates.forEach((date, i) => {
    sheet?.validate(refs[i] ?? 'A1', { date })
  })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(
    back?.validations.map((v) => v.rule),
    dates.map((date) => ({ allowBlank: true, date })),
  )
})

test('a list validation naming no values reads as an empty list', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after:
        '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>""</formula1></dataValidation></dataValidations>',
    }),
  ).sheets[0]
  assert.deepEqual(sheet?.validations, [{ range: 'A1', rule: { allowBlank: false, list: [] } }])
})

test('conditionalFormats round-trip colour scales, cellIs and data bars', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.conditionalFormat('A1:A9', { colorScale: { min: 'FF0000', max: '00FF00' } })
  sheet?.conditionalFormat('B1:B9', { colorScale: { min: 'FF0000', mid: 'FFFF00', max: '00FF00' } })
  sheet?.conditionalFormat('C1', { cellIs: { when: { greaterThan: 5 }, fill: 'FFFF00' } })
  sheet?.conditionalFormat('D1:D9', { dataBar: { color: '638EC6' } })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.conditionalFormats, [
    { range: 'A1:A9', rule: { colorScale: { min: 'FFFF0000', max: 'FF00FF00' } } },
    { range: 'B1:B9', rule: { colorScale: { min: 'FFFF0000', mid: 'FFFFFF00', max: 'FF00FF00' } } },
    { range: 'C1', rule: { cellIs: { when: { greaterThan: 5 }, fill: 'FFFFFF00' } } },
    { range: 'D1:D9', rule: { dataBar: { color: 'FF638EC6' } } },
  ])
})

test('conditionalFormats round-trip expression, duplicates and unique rules', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.conditionalFormat('A1:A9', { expression: { formula: '$A1>0', fill: 'FFFF00' } })
  sheet?.conditionalFormat('B1:B9', { duplicates: { fill: 'FF0000' } })
  sheet?.conditionalFormat('C1:C9', { unique: { fill: '00FF00' } })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.conditionalFormats, [
    { range: 'A1:A9', rule: { expression: { formula: '$A1>0', fill: 'FFFFFF00' } } },
    { range: 'B1:B9', rule: { duplicates: { fill: 'FFFF0000' } } },
    { range: 'C1:C9', rule: { unique: { fill: 'FF00FF00' } } },
  ])
})

test('conditionalFormats round-trip top and bottom rank rules', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.conditionalFormat('A1:A9', { top: { count: 3, fill: 'FFFF00' } })
  sheet?.conditionalFormat('B1:B9', { bottom: { count: 10, fill: 'FF0000', percent: true } })
  const back = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(back?.conditionalFormats, [
    { range: 'A1:A9', rule: { top: { count: 3, fill: 'FFFFFF00' } } },
    { range: 'B1:B9', rule: { bottom: { count: 10, fill: 'FFFF0000', percent: true } } },
  ])
  assert.throws(
    () => sheet?.conditionalFormat('C1', { top: { count: 0, fill: 'FFFFFF' } }),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('conditionalFormats reflect pending and skip kinds not modelled', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after:
        '<conditionalFormatting sqref="A1"><cfRule type="colorScale"><colorScale>' +
        '<cfvo type="min"/><cfvo type="max"/><color theme="4"/><color theme="5"/></colorScale></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1"><cfRule type="colorScale"><colorScale>' +
        '<cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color theme="5"/></colorScale></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="C1"><cfRule type="cellIs" operator="equal" dxfId="99"><formula>1</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="D1"><cfRule type="cellIs" operator="wonky" dxfId="0"><formula>1</formula></cfRule></conditionalFormatting>',
    }),
  ).sheets[0]
  // Theme-only stops, one rgb stop, a missing dxf and an unknown operator are all left out.
  assert.deepEqual(sheet?.conditionalFormats, [])
  sheet?.conditionalFormat('E1:E9', { dataBar: { color: '000000' } })
  assert.deepEqual(sheet?.conditionalFormats, [
    { range: 'E1:E9', rule: { dataBar: { color: 'FF000000' } } },
  ])
})

test('the view getters read the file', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      sheetPr: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
      views:
        '<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" zoomScale="150">' +
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>',
      after: '<autoFilter ref="A1:C1"/>',
    }),
  ).sheets[0]
  assert.equal(sheet?.gridlinesVisible, false)
  assert.equal(sheet?.headingsVisible, false)
  assert.equal(sheet?.zoomPercent, 150)
  assert.equal(sheet?.frozenAt, 'B2')
  assert.equal(sheet?.tabColorHex, 'FFFF0000')
  assert.equal(sheet?.autoFilterRange, 'A1:C1')
})

test('the view getters default sensibly and reflect a pending change', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheets[0]
  assert.equal(sheet?.gridlinesVisible, true)
  assert.equal(sheet?.headingsVisible, true)
  assert.equal(sheet?.zoomPercent, undefined)
  assert.equal(sheet?.frozenAt, undefined)
  assert.equal(sheet?.tabColorHex, undefined)
  assert.equal(sheet?.autoFilterRange, undefined)

  sheet?.showGridlines(false)
  sheet?.zoom(200)
  sheet?.freeze('B2')
  sheet?.tabColor('00FF00')
  sheet?.autoFilter('A1:B2')
  assert.equal(sheet?.gridlinesVisible, false)
  assert.equal(sheet?.zoomPercent, 200)
  assert.equal(sheet?.frozenAt, 'B2')
  assert.equal(sheet?.tabColorHex, 'FF00FF00') // normalised to 8-digit ARGB
  assert.equal(sheet?.autoFilterRange, 'A1:B2')
})

test('rowHeight and columnWidth read the file', () => {
  const sheet = readWorkbook(
    build('<row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row>', {
      views: '<cols><col min="1" max="2" width="12.5" customWidth="1"/></cols>',
    }),
  ).sheets[0]
  assert.equal(sheet?.rowHeight(1), 30)
  assert.equal(sheet?.columnWidth('A'), 12.5)
  assert.equal(sheet?.columnWidth('B'), 12.5) // the same col range covers B
  assert.equal(sheet?.rowHeight(2), undefined)
  assert.equal(sheet?.columnWidth('C'), undefined)
})

test('rowGroupLevel and columnGroupLevel read the file and a pending group', () => {
  const sheet = readWorkbook(
    build('<row r="1" outlineLevel="2"><c r="A1"><v>1</v></c></row>', {
      views: '<cols><col min="2" max="3" outlineLevel="1"/></cols>',
    }),
  ).sheets[0]
  assert.equal(sheet?.rowGroupLevel(1), 2)
  assert.equal(sheet?.rowGroupLevel(2), 0)
  assert.equal(sheet?.columnGroupLevel('B'), 1)
  assert.equal(sheet?.columnGroupLevel('A'), 0)
  sheet?.groupRows(5, 6)
  sheet?.groupColumns('D', 'E')
  assert.equal(sheet?.rowGroupLevel(5), 1)
  assert.equal(sheet?.columnGroupLevel('D'), 1)
})

test('isRowHidden and isColumnHidden read the file and a pending hide', () => {
  const sheet = readWorkbook(
    build('<row r="1" hidden="1"><c r="A1"><v>1</v></c></row>', {
      views: '<cols><col min="2" max="2" hidden="1"/></cols>',
    }),
  ).sheets[0]
  assert.equal(sheet?.isRowHidden(1), true)
  assert.equal(sheet?.isRowHidden(2), false)
  assert.equal(sheet?.isColumnHidden('B'), true)
  assert.equal(sheet?.isColumnHidden('A'), false)
  sheet?.hideRow(2)
  sheet?.hideColumn('A')
  assert.equal(sheet?.isRowHidden(2), true)
  assert.equal(sheet?.isColumnHidden('A'), true)
})

test('rowHeight and columnWidth reflect a pending set before it is written', () => {
  const sheet = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheets[0]
  assert.equal(sheet?.rowHeight(1), undefined)
  sheet?.setRowHeight(1, 42)
  sheet?.setColumnWidth('A', 8)
  assert.equal(sheet?.rowHeight(1), 42)
  assert.equal(sheet?.columnWidth('A'), 8)
})

test('link round-trips: an external url is read back on the cell', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.link('A1', { url: 'https://example.com/a', tooltip: 'go' })
  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')
  assert.deepEqual(cell?.hyperlink, { url: 'https://example.com/a', tooltip: 'go' })
})

test('link round-trips: an internal location is read back on the cell', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.link('A1', { location: 'Sheet2!B2' })
  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')
  assert.deepEqual(cell?.hyperlink, { location: 'Sheet2!B2' })
})

test('a link on a cell with no value is still read', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.link('D4', { url: 'https://x.test/' })
  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('D4')
  assert.deepEqual(cell?.hyperlink, { url: 'https://x.test/' })
  assert.deepEqual(cell?.value, { kind: 'empty' })
})

test('a range hyperlink is reported on its top-left cell', () => {
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after: '<hyperlinks><hyperlink ref="A1:B2" location="Sheet1!C3"/></hyperlinks>',
    }),
  ).sheets[0]
  assert.deepEqual(sheet?.cell('A1')?.hyperlink, { location: 'Sheet1!C3' })
})

test('a cell can carry both a comment and a hyperlink', () => {
  const comments =
    '<comments><authors><author/></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><r><t>note</t></r></text></comment></commentList></comments>'
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after: '<hyperlinks><hyperlink ref="A1" location="Sheet1!Z9"/></hyperlinks>',
      extra: {
        'xl/comments1.xml': comments,
        'xl/worksheets/_rels/sheet1.xml.rels': commentsRels('../comments1.xml'),
      },
    }),
  ).sheets[0]
  const cell = sheet?.cell('A1')
  assert.equal(cell?.comment, 'note')
  assert.deepEqual(cell?.hyperlink, { location: 'Sheet1!Z9' })
})

test('comment writes a comments part, wires it to the sheet and declares its type', () => {
  const contentTypes =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: { '[Content_Types].xml': contentTypes },
    }),
  )
  workbook.sheets[0]?.comment('A1', 'a note')

  const parts = readContainer(workbook.toBytes()).parts
  assert.match(
    decode(parts.get('xl/comments1.xml')),
    /<comment ref="A1" authorId="0"><text><r><t xml:space="preserve">a note<\/t><\/r><\/text><\/comment>/,
  )
  assert.match(
    decode(parts.get('xl/worksheets/_rels/sheet1.xml.rels')),
    /Type="[^"]*relationships\/comments" Target="\.\.\/comments1\.xml"/,
  )
  assert.match(decode(parts.get('[Content_Types].xml')), /PartName="\/xl\/comments1\.xml"/)

  // And it round-trips: reading the written file back sees the comment.
  assert.equal(readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')?.comment, 'a note')
})

test('comment writes a VML drawing and a legacyDrawing so Excel draws the note box', () => {
  const contentTypes =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/></Types>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: { '[Content_Types].xml': contentTypes },
    }),
  )
  workbook.sheets[0]?.comment('A1', 'a note')
  const parts = readContainer(workbook.toBytes()).parts

  const vml = decode(parts.get('xl/drawings/vmlDrawing1.vml'))
  assert.match(vml, /ObjectType="Note"/)
  assert.match(vml, /<x:Row>0<\/x:Row><x:Column>0<\/x:Column>/)

  // The comment took rId1, so the drawing takes rId2, and the sheet points at it.
  assert.match(
    decode(parts.get('xl/worksheets/sheet1.xml')),
    /<legacyDrawing r:id="rId2"\/><\/worksheet>/,
  )
  assert.match(
    decode(parts.get('xl/worksheets/_rels/sheet1.xml.rels')),
    /Id="rId2"[^>]*relationships\/vmlDrawing[^>]*Target="\.\.\/drawings\/vmlDrawing1\.vml"/,
  )
  assert.match(
    decode(parts.get('[Content_Types].xml')),
    /PartName="\/xl\/drawings\/vmlDrawing1\.vml"/,
  )
})

test('comment inserts legacyDrawing before tableParts in schema order', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after: '<tableParts count="1"><tablePart r:id="rId9"/></tableParts>',
    }),
  )
  workbook.sheets[0]?.comment('A1', 'note')
  assert.match(
    decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml')),
    /<legacyDrawing r:id="rId2"\/><tableParts/,
  )
})

test('comment inserts legacyDrawing before a worksheet extLst', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      after: '<extLst><ext uri="{FEATURE}"><x14:sparklineGroups/></ext></extLst>',
    }),
  )
  workbook.sheets[0]?.comment('A1', 'note')
  assert.match(
    decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml')),
    /<legacyDrawing r:id="rId2"\/><extLst>/,
  )
})

test('comment joins an existing sheet rels part rather than replacing it', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://x/drawing" Target="../drawings/d1.xml"/></Relationships>',
      },
    }),
  )
  workbook.sheets[0]?.comment('A1', 'note')

  const rels = decode(
    readContainer(workbook.toBytes()).parts.get('xl/worksheets/_rels/sheet1.xml.rels'),
  )
  assert.match(rels, /Target="\.\.\/drawings\/d1\.xml"/) // the one it had is kept
  assert.match(rels, /relationships\/comments/) // and the comment is added
})

test('comment refuses a cell that already carries a comment, not the whole sheet', () => {
  const comments =
    '<comments><authors><author/></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><t>existing</t></text></comment></commentList></comments>'
  const sheet = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/comments1.xml': comments,
        'xl/worksheets/_rels/sheet1.xml.rels': commentsRels('../comments1.xml'),
      },
    }),
  ).sheet('Data')

  assert.throws(
    () => sheet?.comment('A1', 'new'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unsupported-edit',
  )
  // A different cell is fine.
  assert.doesNotThrow(() => sheet?.comment('B2', 'added'))
})

test('comment on a sheet with comments splices in without rebuilding the rich text', () => {
  const comments =
    '<comments><authors><author>Ada</author></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>rich existing</t></r></text></comment>' +
    '</commentList></comments>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="C3"><v>2</v></c></row>', {
      extra: {
        'xl/comments1.xml': comments,
        // A rels part naming the comments part and an existing VML drawing (B1).
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="${COMMENTS_REL}" Target="../comments1.xml"/>` +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
          '</Relationships>',
        'xl/drawings/vmlDrawing1.vml': buildVmlDrawing(['A1']),
      },
    }),
  )
  workbook.sheet('Data')?.comment('C3', 'appended')
  const parts = readContainer(workbook.toBytes()).parts

  const merged = decode(parts.get('xl/comments1.xml'))
  assert.match(merged, /<r><rPr><b\/><\/rPr><t>rich existing<\/t><\/r>/) // untouched
  assert.match(merged, /<comment ref="C3" authorId="0">/) // added
  const vml = decode(parts.get('xl/drawings/vmlDrawing1.vml'))
  assert.match(vml, /<v:shape id="_x0000_s1025"/) // the one it had
  assert.match(vml, /<v:shape id="_x0000_s1026"/) // and the appended note
  // Both comments read back.
  const back = readWorkbook(workbook.toBytes()).sheet('Data')
  assert.equal(back?.cell('A1')?.comment, 'rich existing')
  assert.equal(back?.cell('C3')?.comment, 'appended')
})

test('comment on a sheet whose comments lack a drawing authors one', () => {
  const comments =
    '<comments><authors><author/></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><t>plain</t></text></comment></commentList></comments>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>', {
      extra: {
        'xl/comments1.xml': comments,
        'xl/worksheets/_rels/sheet1.xml.rels': commentsRels('../comments1.xml'), // no VML (B2)
      },
    }),
  )
  workbook.sheet('Data')?.comment('B2', 'added')
  const parts = readContainer(workbook.toBytes()).parts

  // A drawing is authored and the sheet is pointed at it.
  assert.match(decode(parts.get('xl/drawings/vmlDrawing1.vml')), /ObjectType="Note"/)
  assert.match(decode(parts.get('xl/worksheets/sheet1.xml')), /<legacyDrawing r:id="/)
  assert.equal(readWorkbook(workbook.toBytes()).sheet('Data')?.cell('B2')?.comment, 'added')
})

test('set applies an alignment, adding it to the cell format', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { alignment: { horizontal: 'center', wrapText: true } })

  const styles = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))

  assert.match(styles, /<alignment horizontal="center" wrapText="1"\/>/)
  assert.match(lastXf(styles), /applyAlignment="1"/)
})

test('format applies an alignment without touching the value', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>7</v></c></row>'))
  workbook.sheets[0]?.format('A1', { alignment: { vertical: 'top' } })

  const parts = readContainer(workbook.toBytes()).parts

  assert.match(decode(parts.get('xl/styles.xml')), /<alignment vertical="top"\/>/)
  assert.match(decode(parts.get('xl/worksheets/sheet1.xml')), /<c [^>]*s="\d+"[^>]*><v>7<\/v><\/c>/)
})

test('set composes a font and an alignment into one cell format', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { font: { bold: true }, alignment: { horizontal: 'right' } })

  const xf = lastXf(decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml')))

  assert.match(xf, /applyFont="1"/)
  assert.match(xf, /applyAlignment="1"/)
  assert.match(xf, /<alignment horizontal="right"\/>/)
})

test('set refuses a bad text rotation, before recording anything', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.throws(() => sheet?.set('A1', 'x', { alignment: { textRotation: 999 } }), /rotation/)
  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
})

test('strikethrough, a superscript and an underline style read back off a cell', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', {
    font: { strike: true, verticalAlign: 'superscript', underline: 'double' },
  })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.font, {
    strike: true,
    verticalAlign: 'superscript',
    underline: 'double',
  })
})

test('an alignment set on a cell reads back off it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { alignment: { horizontal: 'center', wrapText: true } })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.alignment, {
    horizontal: 'center',
    wrapText: true,
  })
})

test('protection set on a cell reads back off it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { protection: { locked: false, hidden: true } })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.protection, { locked: false, hidden: true })
})

test('format unlocks a cell without touching its value', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>7</v></c></row>'))
  workbook.sheets[0]?.format('A1', { protection: { locked: false } })

  const parts = readContainer(workbook.toBytes()).parts
  assert.match(decode(parts.get('xl/styles.xml')), /<protection locked="0"\/>/)
  assert.match(decode(parts.get('xl/worksheets/sheet1.xml')), /<c [^>]*s="\d+"[^>]*><v>7<\/v><\/c>/)
})

test('protect turns on sheet protection with the Excel defaults, after sheetData', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.protect()

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<\/sheetData><sheetProtection sheet="1" objects="1" scenarios="1"\/>/)
})

test('protect permits chosen actions and can bar selecting locked cells', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.protect({ formatCells: true, sort: true, selectLockedCells: false })

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /formatCells="0"/)
  assert.match(sheet, /sort="0"/)
  assert.match(sheet, /selectLockedCells="1"/)
})

test('protecting a sheet with no other edit still rewrites only that sheet', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.protect()

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<sheetProtection /)
})

test('unprotect removes the protection a file already declared', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/sheet1.xml':
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
          '<sheetProtection sheet="1"/></worksheet>',
      },
    }),
  )
  workbook.sheets[0]?.unprotect()

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.doesNotMatch(sheet, /sheetProtection/)
})

test('setRowHeight sizes a row as a custom height', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.setRowHeight(1, 42)

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<row customHeight="1" ht="42" r="1">/)
})

test('setRowHeight refuses a row below one or a negative height', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.throws(
    () => sheet?.setRowHeight(0, 20),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => sheet?.setRowHeight(1, -5),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('setColumnWidth opens a cols element for the column', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.setColumnWidth('B', 24)

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<cols><col min="2" max="2" width="24" customWidth="1"\/><\/cols><sheetData>/)
})

test('setColumnWidth refuses a bad column or a negative width', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.throws(
    () => sheet?.setColumnWidth('7', 10),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => sheet?.setColumnWidth('A', -1),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('setRowHeight refuses a row past the sheet maximum', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(
    () => workbook.sheets[0]?.setRowHeight(1048577, 20),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('setColumnWidth refuses a column past the last', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(
    () => workbook.sheets[0]?.setColumnWidth('XFE', 15),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('protection reads back off a protected sheet', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/sheet1.xml':
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
          '<sheetProtection sheet="1" objects="1" scenarios="1" sort="0"/></worksheet>',
      },
    }),
  )

  assert.deepEqual(workbook.sheets[0]?.protection, {
    editObjects: false,
    editScenarios: false,
    sort: true,
  })
})

test('a plain sheet reports no protection', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.equal(workbook.sheets[0]?.protection, undefined)
})

test('protection reflects a pending protect and unprotect', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  sheet?.protect({ sort: true })
  assert.deepEqual(sheet?.protection, { sort: true })

  sheet?.unprotect()
  assert.equal(sheet?.protection, undefined)
})

test('merge adds a mergeCells element after sheetData', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.merge('A1:B2')

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<\/sheetData><mergeCells count="1"><mergeCell ref="A1:B2"\/><\/mergeCells>/)
})

test('merge canonicalises the range and refuses one without a colon', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.merge('b2:a1')

  assert.match(
    decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml')),
    /<mergeCell ref="A1:B2"\/>/,
  )
  assert.throws(
    () => sheet?.merge('A1'),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
})

test('merge refuses a range whose cells are outside the sheet, the way set does', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  const refuses = (range: string) =>
    assert.throws(
      () => sheet?.merge(range),
      (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
    )
  refuses('A0:B2') // row zero, which set() also refuses
  refuses('A1:B1048577') // past the last row a sheet can hold
})

test('a value written into a merged non-anchor cell is still refused', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.merge('A1:B2')

  const reopened = readWorkbook(workbook.toBytes())
  assert.throws(
    () => reopened.sheets[0]?.set('B2', 5),
    (error: unknown) => error instanceof XlsxError && /merged into A1/.test(error.message),
  )
})

test('a plain value writes into a package with no style table', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
      [
        'xl/worksheets/sheet1.xml',
        encode(
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        ),
      ],
    ]),
  })
  const workbook = readWorkbook(bytes)
  workbook.sheets[0]?.set('A1', 42)

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<c r="A1"><v>42<\/v><\/c>/)
})

test('format refuses when the package has no style table', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
      [
        'xl/worksheets/sheet1.xml',
        encode(
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        ),
      ],
    ]),
  })

  assert.throws(
    () => readWorkbook(bytes).sheets[0]?.format('A1', { border: { all: { style: 'thin' } } }),
    /no style table/,
  )
})

test('format keeps a cell whose number format id resolves to nothing', () => {
  // s=1 points at numFmtId 200, which no numFmts entry defines, so the cell has
  // no resolvable number format; restyling with a font must not invent one.
  const styles =
    '<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="200"/></cellXfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="1"><v>5</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )
  workbook.sheets[0]?.format('A1', { font: { bold: true } })

  assert.equal(workbook.sheets[0]?.cell('A1')?.numberFormat, undefined)
})

test('format applies a border, merging onto the sides the cell has', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.format('A1', { border: { all: { style: 'thin', color: '000000' } } })

  const parts = readContainer(workbook.toBytes()).parts
  const styles = decode(parts.get('xl/styles.xml'))
  const sheet = decode(parts.get('xl/worksheets/sheet1.xml'))

  assert.match(styles, /<border><left style="thin"><color rgb="FF000000"\/><\/left>/)
  assert.match(sheet, /<c s="\d+" r="A1"><v>1<\/v><\/c>/)
})

test('cell() reads back the font, fill and border it was given', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', {
    font: { bold: true },
    fill: { type: 'solid', color: 'FF0000' },
    border: { top: { style: 'thin' } },
  })

  const cell = workbook.sheets[0]?.cell('A1')
  assert.deepEqual(cell?.font, { bold: true })
  assert.deepEqual(cell?.fill, { type: 'solid', color: 'FFFF0000' })
  assert.deepEqual(cell?.border, { top: { style: 'thin' } })
})

test('a diagonal border is written and read back', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', {
    border: { diagonal: { style: 'thin', color: 'FF0000', down: true } },
  })

  const styles = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(
    styles,
    /<border diagonalDown="1"><left\/><right\/><top\/><bottom\/><diagonal style="thin"><color rgb="FFFF0000"\/><\/diagonal><\/border>/,
  )
  assert.deepEqual(workbook.sheets[0]?.cell('A1')?.border, {
    diagonal: { style: 'thin', color: 'FFFF0000', down: true },
  })
})

test('a diagonal-up border sets diagonalUp and reads back', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { border: { diagonal: { style: 'thin', up: true } } })

  const styles = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(styles, /<border diagonalUp="1">.*<diagonal style="thin"\/><\/border>/)
  assert.deepEqual(workbook.sheets[0]?.cell('A1')?.border, {
    diagonal: { style: 'thin', up: true },
  })
})

test('cell() reads formatting a file already carries', () => {
  const styles =
    '<styleSheet><fonts count="2"><font/><font><b/><sz val="12"/></font></fonts>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><bottom style="medium"/><left/><right/><top/><diagonal/></border></borders>' +
    '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0" fontId="1" borderId="1"/></cellXfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )

  const cell = workbook.sheets[0]?.cell('A1')
  assert.deepEqual(cell?.font, { bold: true, size: 12 })
  assert.deepEqual(cell?.border, { bottom: { style: 'medium' } })
  assert.equal(cell?.fill, undefined)
})

test('a plain cell reports no font, fill or border', () => {
  const cell = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>')).sheets[0]?.cell('A1')

  assert.equal(cell?.font, undefined)
  assert.equal(cell?.fill, undefined)
  assert.equal(cell?.border, undefined)
})

test('set and format apply a solid fill, composing with a font', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', {
    fill: { type: 'solid', color: '00FF00' },
    font: { bold: true },
  })

  const styles = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(styles, /<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"\//)
  const xf = [...styles.matchAll(/<xf [^>]*\/>/g)].map((match) => match[0]).at(-1) ?? ''
  assert.match(xf, /applyFill="1"/)
  assert.match(xf, /applyFont="1"/)
})

test('a pattern fill set on a cell reads back off it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', {
    fill: { type: 'pattern', pattern: 'lightGrid', color: 'FF0000', background: 'FFFFFF' },
  })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.fill, {
    type: 'pattern',
    pattern: 'lightGrid',
    color: 'FFFF0000',
    background: 'FFFFFFFF',
  })
})

test('a theme colour set on a cell font reads back off it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'x', { font: { color: { theme: 4, tint: 0.4 } } })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.font, { color: { theme: 4, tint: 0.4 } })
})

test('editing a cell keeps the theme colour its font already carries', () => {
  const styles =
    '<styleSheet><fonts count="2"><font/><font><color theme="1"/><sz val="12"/></font></fonts>' +
    '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0" fontId="1"/></cellXfs></styleSheet>'
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>', { extra: { 'xl/styles.xml': styles } }),
  )
  workbook.sheets[0]?.set('A1', 'x', { font: { bold: true } })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual(reopened.sheets[0]?.cell('A1')?.font, {
    bold: true,
    color: { theme: 1 },
    size: 12,
  })
})

test('format restyles a formula cell without touching its formula', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row>'))
  workbook.sheets[0]?.format('A1', { font: { bold: true } })

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<c s="\d+" r="A1"><f>SUM\(B1:B2\)<\/f><v>3<\/v><\/c>/)
})

test('format applies a number format and cell() reflects it at once', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>45292</v></c></row>'))
  workbook.sheets[0]?.format('A1', { numberFormat: 'yyyy-mm-dd' })

  assert.equal(workbook.sheets[0]?.cell('A1')?.value.kind, 'date')
})

test('format that removes a date format turns the cell back into a number', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'))
  assert.equal(workbook.sheets[0]?.cell('A1')?.value.kind, 'date')

  workbook.sheets[0]?.format('A1', { numberFormat: '0.00' })

  assert.equal(workbook.sheets[0]?.cell('A1')?.value.kind, 'number')
})

test('format creates an empty styled cell when the cell is not there', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.format('B1', { font: { italic: true } })

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<c r="B1" s="\d+"\/>/)
})

test('format is allowed on a shared-formula master, keeping the formula', () => {
  const workbook = readWorkbook(
    build(
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>2</v></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c></row>',
    ),
  )
  assert.doesNotThrow(() => workbook.sheets[0]?.format('A1', { font: { bold: true } }))

  const sheet = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<c s="\d+" r="A1"><f t="shared" ref="A1:A2" si="0">B1<\/f><v>2<\/v><\/c>/)
})

test('format with empty options is a no-op', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.format('A1', {})

  assert.deepEqual(workbook.sheets[0]?.cell('A1')?.value, { kind: 'number', value: 1 })
})

test('puts new text in the shared string table when the file has one', () => {
  const workbook = readWorkbook(
    withStrings(
      '<row r="1"><c r="A1"><v>1</v></c></row>',
      '<sst count="1" uniqueCount="1"><si><t>existing</t></si></sst>',
    ),
  )
  workbook.sheets[0]?.set('A1', 'fresh')

  const parts = readContainer(workbook.toBytes()).parts
  const table = new TextDecoder().decode(parts.get('xl/sharedStrings.xml') ?? new Uint8Array())
  const sheet = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml') ?? new Uint8Array())

  assert.match(table, /<si><t>fresh<\/t><\/si>/)
  assert.match(sheet, /<c r="A1" t="s"><v>1<\/v><\/c>/)
})

test('reuses a string the table already holds', () => {
  const workbook = readWorkbook(
    withStrings(
      '<row r="1"><c r="A1"><v>1</v></c></row>',
      '<sst count="1" uniqueCount="1"><si><t>existing</t></si></sst>',
    ),
  )
  workbook.sheets[0]?.set('A1', 'existing')

  const parts = readContainer(workbook.toBytes()).parts
  const table = new TextDecoder().decode(parts.get('xl/sharedStrings.xml') ?? new Uint8Array())

  assert.equal(table.match(/<si>/g)?.length, 1)
  assert.match(table, /uniqueCount="1"/)
})

test('stores repeated text once no matter how many cells use it', () => {
  const workbook = readWorkbook(
    withStrings(
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
      '<sst count="0" uniqueCount="0"></sst>',
    ),
  )
  workbook.sheets[0]?.set('A1', 'repeated')
  workbook.sheets[0]?.set('B1', 'repeated')

  const parts = readContainer(workbook.toBytes()).parts
  const table = new TextDecoder().decode(parts.get('xl/sharedStrings.xml') ?? new Uint8Array())

  assert.equal(table.match(/<si>/g)?.length, 1)
})

test('reads back text written through the shared string table', () => {
  const workbook = readWorkbook(
    withStrings('<row r="1"><c r="A1"><v>1</v></c></row>', '<sst count="0" uniqueCount="0"></sst>'),
  )
  workbook.sheets[0]?.set('A1', 'round tripped')

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'text', value: 'round tripped' })
})

test('falls back to an inline string when the file has no table', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'inline')

  const parts = readContainer(workbook.toBytes()).parts
  const sheet = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml') ?? new Uint8Array())

  assert.equal(parts.has('xl/sharedStrings.xml'), false)
  assert.match(sheet, /t="inlineStr"><is><t>inline<\/t><\/is>/)
})

test('a date written into an unformatted cell reads back as a date', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', new Date(2024, 2, 5))

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind, 'date')
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getMonth(), 2)
  assert.equal(cell?.value.kind === 'date' && cell.value.value.getDate(), 5)
})

test('a date written into a new cell reads back as a date', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('D4', new Date(2024, 2, 5))

  const reopened = readWorkbook(workbook.toBytes())
  const cell = [...(reopened.sheets[0]?.cells() ?? [])].find((each) => each.reference === 'D4')

  assert.equal(cell?.value.kind, 'date')
})

test('a cell that already has a date format keeps it', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', new Date(2024, 2, 5))

  const parts = readContainer(workbook.toBytes()).parts
  const styles = new TextDecoder().decode(parts.get('xl/styles.xml') ?? new Uint8Array())

  assert.equal(styles.match(/<xf /g)?.length, 2)
})

test('reuses a date format the file already has', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'),
  )
  workbook.sheets[0]?.set('A1', new Date(2024, 2, 5))
  workbook.sheets[0]?.set('B1', new Date(2024, 3, 6))

  const parts = readContainer(workbook.toBytes()).parts
  const styles = new TextDecoder().decode(parts.get('xl/styles.xml') ?? new Uint8Array())

  assert.equal(styles.match(/<xf /g)?.length, 2)
  for (const cell of readWorkbook(workbook.toBytes()).sheets[0]?.cells() ?? []) {
    assert.equal(cell.value.kind, 'date', `${cell.reference} is not a date`)
  }
})

test('adds one date format when the file has none, however many dates are written', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>', {
      extra: {
        'xl/styles.xml': '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', new Date(2024, 2, 5))
  workbook.sheets[0]?.set('B1', new Date(2024, 3, 6))

  const parts = readContainer(workbook.toBytes()).parts
  const styles = new TextDecoder().decode(parts.get('xl/styles.xml') ?? new Uint8Array())

  assert.equal(styles.match(/<xf /g)?.length, 2)
  for (const cell of readWorkbook(workbook.toBytes()).sheets[0]?.cells() ?? []) {
    assert.equal(cell.value.kind, 'date', `${cell.reference} is not a date`)
  }
})

test('a number written to a cell does not gain a date format', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 42)

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'number', value: 42 })
})

test('finds a sheet by name', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.equal(workbook.sheet('Data')?.name, 'Data')
})

test('reports no sheet when the name is not there', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.equal(workbook.sheet('Missing'), undefined)
})

test('reads a single cell by reference', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>'),
  )

  assert.deepEqual(workbook.sheets[0]?.cell('C1')?.value, { kind: 'number', value: 3 })
})

test('reports no cell when nothing is there', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.equal(workbook.sheets[0]?.cell('Z9'), undefined)
})

test('reads a single cell that was just set', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('B2', 'fresh')

  assert.deepEqual(workbook.sheets[0]?.cell('B2')?.value, { kind: 'text', value: 'fresh' })
})

test('normalises the reference it is asked for', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.equal(workbook.sheets[0]?.cell('$a$1')?.reference, 'A1')
})

test('a cleared cell is still visited, as empty', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', null)

  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'empty' })
})

test('drops the calculation chain when a cell is written', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row>', {
      extra: {
        'xl/calcChain.xml': '<calcChain><c r="A1" i="1"/></calcChain>',
        '[Content_Types].xml':
          '<Types><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/><Override PartName="/xl/workbook.xml" ContentType="x"/></Types>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', 5)

  const parts = readContainer(workbook.toBytes()).parts
  const types = new TextDecoder().decode(parts.get('[Content_Types].xml') ?? new Uint8Array())

  assert.equal(parts.has('xl/calcChain.xml'), false)
  assert.equal(types.includes('calcChain'), false, 'the content type override was left behind')
  assert.match(types, /workbook\.xml/, 'the other overrides were lost')
})

test('leaves the calculation chain alone when nothing is written', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: { 'xl/calcChain.xml': '<calcChain/>' },
    }),
  )

  assert.equal(readContainer(workbook.toBytes()).parts.has('xl/calcChain.xml'), true)
})

test('reports a part that is not valid utf-8 rather than mangling it', () => {
  const parts = new Map([
    ['_rels/.rels', encode(ROOT_RELS)],
    ['xl/workbook.xml', new Uint8Array([0x3c, 0xff, 0xfe, 0x3e])],
  ])

  assert.throws(() => readWorkbook(writeContainer({ parts })), /xl\/workbook\.xml/)
})

test('reads one cell without reparsing the sheet for each lookup', () => {
  const size = 2000
  const rows = Array.from(
    { length: size },
    (_unused, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`,
  ).join('')
  const workbook = readWorkbook(build(rows))
  const sheet = workbook.sheets[0]

  const started = process.hrtime.bigint()
  for (let row = 1; row <= size; row++) sheet?.cell(`A${row}`)
  const elapsedMs = Number(process.hrtime.bigint() - started) / 1e6

  // Reparsing per lookup takes about 1.9s here; an index takes a few ms.
  assert.ok(elapsedMs < 500, `${size} lookups took ${elapsedMs.toFixed(0)}ms`)
})

test('a lookup sees a value written after the last lookup', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
  sheet?.set('A1', 2)
  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 2 })
})

test('reports which part is not valid utf-8 when it is not the workbook', () => {
  const parts = new Map([
    ['_rels/.rels', encode(ROOT_RELS)],
    [
      'xl/workbook.xml',
      encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
    ],
    ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
    ['xl/sharedStrings.xml', new Uint8Array([0x3c, 0xff, 0xfe, 0x3e])],
  ])

  assert.throws(() => readWorkbook(writeContainer({ parts })), /sharedStrings\.xml is not valid/)
})

test('an error carries a code and the part it came from', () => {
  const parts = new Map([['xl/workbook.xml', encode('<workbook/>')]])

  try {
    readWorkbook(writeContainer({ parts }))
    assert.fail('expected a failure')
  } catch (error) {
    assert.ok(error instanceof XlsxError)
    assert.equal(error.code, 'missing-part')
    assert.equal(error.part, '_rels/.rels')
  }
})

test('a bad reference names itself', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  try {
    workbook.sheets[0]?.set('nonsense', 1)
    assert.fail('expected a failure')
  } catch (error) {
    assert.ok(error instanceof XlsxError)
    assert.equal(error.code, 'bad-reference')
    assert.equal(error.reference, 'nonsense')
  }
})

test('asks for a recalculation once a formula is written', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.match(book, /<calcPr[^>]*fullCalcOnLoad="1"/)
})

test('sets the recalculation flag on an existing calcPr without losing it', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '<calcPr calcId="140000" iterate="1"/></workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.match(book, /iterate="1"/)
  assert.match(book, /fullCalcOnLoad="1"/)
  assert.equal(book.match(/<calcPr/g)?.length, 1)
})

test('leaves the workbook part alone when no formula is written', () => {
  const original = build('<row r="1"><c r="A1"><v>1</v></c></row>')
  const workbook = readWorkbook(original)
  workbook.sheets[0]?.set('A1', 5)

  const before = readContainer(original).parts.get('xl/workbook.xml')
  const after = readContainer(workbook.toBytes()).parts.get('xl/workbook.xml')

  assert.deepEqual(after, before)
})

test('turns an existing recalculation flag back on', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '<calcPr fullCalcOnLoad="0"/></workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.match(book, /fullCalcOnLoad="1"/)
  assert.equal(book.includes('fullCalcOnLoad="0"'), false)
})

test('adds the flag to a calcPr written with a closing tag', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '<calcPr calcId="1"></calcPr></workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.match(book, /<calcPr calcId="1" fullCalcOnLoad="1"><\/calcPr>/)
})

test('adds a recalculation flag to a prefixed workbook', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          '<x:workbook><x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.match(book, /<x:calcPr fullCalcOnLoad="1"\/><\/x:workbook>/)
})

test('writes a number with the format that was asked for', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 1234.5, { numberFormat: '"$"#,##0.00' })

  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')

  assert.deepEqual(cell?.value, { kind: 'number', value: 1234.5 })
  assert.equal(cell?.numberFormat, '"$"#,##0.00')
})

test('two cells asking for one format share a single style', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'),
  )
  workbook.sheets[0]?.set('A1', 1, { numberFormat: '0.0%' })
  workbook.sheets[0]?.set('B1', 2, { numberFormat: '0.0%' })

  const parts = readContainer(workbook.toBytes()).parts
  const styles = new TextDecoder().decode(parts.get('xl/styles.xml') ?? new Uint8Array())
  const sheet = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml') ?? new Uint8Array())

  assert.equal(styles.match(/<xf /g)?.length, 3)
  const used = [...sheet.matchAll(/<c r="[AB]1" s="(\d+)"/g)].map((m) => m[1])
  assert.deepEqual(used, [used[0], used[0]])
})

test('a format overrides the date format a Date would otherwise get', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', new Date(2024, 0, 1), { numberFormat: 'yyyy' })

  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')

  assert.equal(cell?.numberFormat, 'yyyy')
  assert.equal(cell?.value.kind, 'date')
})

test('a cell written with no format keeps the one it had', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 5)

  const cell = readWorkbook(workbook.toBytes()).sheets[0]?.cell('A1')

  assert.equal(cell?.numberFormat, 'yyyy-mm-dd')
})

test('reads taken between edits agree with a full read', () => {
  const rows = Array.from(
    { length: 30 },
    (_unused, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`,
  ).join('')
  const workbook = readWorkbook(build(rows))
  const sheet = workbook.sheets[0]

  // The first read builds the index, so the writes after it are predicted
  // rather than sent back to the sheet.
  assert.equal(sheet?.cell('A1')?.value.kind, 'number')

  const writes: Array<[string, CellInput, string | undefined]> = [
    ['A1', 'text', undefined],
    ['A2', 42, undefined],
    ['A3', true, undefined],
    ['A4', null, undefined],
    ['A6', { formula: 'A1+1' }, undefined],
    ['A7', 9, '0.0%'],
    ['A8', new Date(2024, 0, 1), 'yyyy-mm-dd'],
    ['A9', 3, '#,##0'],
    ['Z40', 'new', undefined],
  ]

  for (const [reference, value, format] of writes) {
    sheet?.set(reference, value, format === undefined ? undefined : { numberFormat: format })
    assert.equal(sheet?.cell(reference)?.numberFormat, format, `${reference} format`)
  }

  // A date written over a cell that already shows dates keeps that format.
  sheet?.set('A8', new Date(2024, 5, 6))
  assert.equal(sheet?.cell('A8')?.numberFormat, 'yyyy-mm-dd')

  for (const cell of sheet?.cells() ?? []) {
    assert.deepEqual(
      sheet?.cell(cell.reference)?.value,
      cell.value,
      `${cell.reference} read one way but not the other`,
    )
    assert.deepEqual(sheet?.cell(cell.reference)?.formula, cell.formula, cell.reference)
    assert.equal(sheet?.cell(cell.reference)?.numberFormat, cell.numberFormat, cell.reference)
  }
})

test('a date written onto a plain cell still reads the same both ways', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.equal(sheet?.cell('A1')?.value.kind, 'number')
  sheet?.set('A1', new Date(2024, 0, 1))

  const [cell] = [...(sheet?.cells() ?? [])]
  assert.deepEqual(sheet?.cell('A1')?.value, cell?.value)
})

test('a single cell read agrees with a full read after many edits', () => {
  const rows = Array.from(
    { length: 30 },
    (_unused, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`,
  ).join('')
  const workbook = readWorkbook(build(rows))
  const sheet = workbook.sheets[0]

  sheet?.set('A1', 'text')
  sheet?.set('A2', 42)
  sheet?.set('A3', true)
  sheet?.set('A4', null)
  sheet?.set('A5', new Date(2024, 0, 1))
  sheet?.set('A6', { formula: 'A1+1' })
  sheet?.set('B7', 9, { numberFormat: '0.0%' })
  sheet?.set('Z40', 'new')

  // cells() re-reads the patched sheet, so it is the authority.
  for (const cell of sheet?.cells() ?? []) {
    assert.deepEqual(
      sheet?.cell(cell.reference)?.value,
      cell.value,
      `${cell.reference} read one way but not the other`,
    )
    assert.deepEqual(sheet?.cell(cell.reference)?.formula, cell.formula, cell.reference)
    assert.equal(sheet?.cell(cell.reference)?.numberFormat, cell.numberFormat, cell.reference)
  }
})

test('leaves chartsheets out of the sheets a caller can edit', async () => {
  const bytes = new Uint8Array(await readFile('fixtures/real/WithChartSheet.xlsx'))
  const workbook = readWorkbook(bytes)

  // A chartsheet holds no cells, so exposing it as a Worksheet means cells()
  // reports an empty sheet and set() reports the file is malformed.
  assert.equal(
    workbook.sheets.some((sheet) => sheet.name === 'Chart2'),
    false,
  )

  // Still written back untouched, like every other part we do not interpret.
  const before = readContainer(bytes)
  const after = readContainer(workbook.toBytes())
  for (const path of before.parts.keys()) {
    assert.equal(after.parts.has(path), true, `lost ${path}`)
  }
  assert.equal(after.parts.has('xl/chartsheets/sheet1.xml'), true)
})

test('a carriage return survives a round trip through the file', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A1', 'a\r\nb')

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'text', value: 'a\r\nb' })
})

test('marks a workbook for recalculation whichever quotes it used', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          "<workbook><calcPr calcId='1' fullCalcOnLoad='0'/><sheets>" +
          '<sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const book = new TextDecoder().decode(
    readContainer(workbook.toBytes()).parts.get('xl/workbook.xml') ?? new Uint8Array(),
  )

  assert.match(book, /fullCalcOnLoad=("|')1\1/)
})

const REFUSED: ReadonlyArray<readonly [string, CellInput, RegExp]> = [
  ['a value that is not a number', Number.NaN, /A1 cannot hold NaN/],
  ['an infinite number', Number.POSITIVE_INFINITY, /A1 cannot hold Infinity/],
  ['text xml cannot hold', 'a\u0000b', /A1 holds U\+0000/],
  ['a formula xml cannot hold', { formula: 'LEN("\u0000")' }, /A1 holds U\+0000/],
  ['a date before the epoch', new Date(1800, 0, 1), /to cell A1/],
  ['an invalid date', new Date('nonsense'), /to cell A1/],
]

for (const [what, value, message] of REFUSED) {
  test(`refuses ${what} at the set() call, not at save time`, () => {
    const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
    const sheet = workbook.sheets[0]

    assert.throws(
      () => sheet?.set('A1', value),
      (error: unknown) => {
        assert.ok(error instanceof XlsxError)
        assert.equal(error.code, 'unwritable-value')
        assert.equal(error.reference, 'A1')
        assert.match(error.message, message)
        return true
      },
    )

    // The refused edit was never recorded, so nothing downstream sees it.
    assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
    assert.deepEqual(
      [...(sheet?.cells() ?? [])].map((cell) => cell.value),
      [{ kind: 'number', value: 1 }],
    )
    const reopened = readWorkbook(workbook.toBytes())
    assert.deepEqual(reopened.sheets[0]?.cell('A1')?.value, { kind: 'number', value: 1 })
  })
}

test('refuses to overwrite a shared formula master at the set() call', () => {
  const workbook = readWorkbook(
    build(
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>2</v></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>',
    ),
  )
  const sheet = workbook.sheets[0]

  assert.throws(() => sheet?.set('A1', 5), /A1 defines shared formula 0/)

  assert.deepEqual(sheet?.cell('A1')?.formula, { kind: 'expression', expression: 'B1*2' })
  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 2 })
})

test('a refused edit leaves the rest of the batch writable', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  sheet?.set('A2', 'kept')
  assert.throws(() => sheet?.set('A3', Number.NaN), /cannot hold/)
  sheet?.set('A4', 'also kept')

  const reopened = readWorkbook(workbook.toBytes()).sheets[0]
  assert.deepEqual(reopened?.cell('A2')?.value, { kind: 'text', value: 'kept' })
  assert.equal(reopened?.cell('A3'), undefined)
  assert.deepEqual(reopened?.cell('A4')?.value, { kind: 'text', value: 'also kept' })
})

test('reads a large number in a date formatted cell as the number it is', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1" s="1"><v>3000000</v></c></row>'))

  const [cell] = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(cell?.value, { kind: 'number', value: 3000000 })
})

test('refuses a date the workbook cannot hold, naming the cell', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(
    () => workbook.sheets[0]?.set('A1', new Date(50000, 0, 1)),
    (error: unknown) => error instanceof XlsxError && /A1/.test(error.message),
  )
})

test('records nothing when a write is refused after the value is checked', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]

  assert.throws(() => sheet?.set('A1', 5, { numberFormat: `x${String.fromCharCode(0)}y` }))

  // A refused edit leaves the workbook exactly as it was: readable, and with
  // nothing queued for the file it is about to write.
  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
  assert.deepEqual([...(sheet?.cells() ?? [])][0]?.value, { kind: 'number', value: 1 })

  const reopened = readWorkbook(workbook.toBytes())
  assert.deepEqual([...(reopened.sheets[0]?.cells() ?? [])][0]?.value, { kind: 'number', value: 1 })
})

test('drops the calculation chain relationship along with the part', () => {
  const withChain = build('<row r="1"><c r="A1"><v>1</v></c></row>', {
    extra: {
      'xl/calcChain.xml': '<calcChain><c r="A1"/></calcChain>',
      'xl/_rels/workbook.xml.rels':
        '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>' +
        '</Relationships>',
    },
  })

  const workbook = readWorkbook(withChain)
  workbook.sheets[0]?.set('A1', 2)
  const parts = readContainer(workbook.toBytes()).parts

  assert.equal(parts.has('xl/calcChain.xml'), false)
  // A relationship pointing at a part that is gone is an invalid package, and
  // it is the kind of thing LibreOffice ignores and Excel offers to repair.
  const rels = new TextDecoder().decode(parts.get('xl/_rels/workbook.xml.rels'))
  assert.doesNotMatch(rels, /calcChain/)
  assert.match(rels, /worksheets\/sheet1\.xml/)
})

test('refuses a write to a sheet whose part is not in the package', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="Gone" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
    ]),
  })
  const workbook = readWorkbook(bytes)

  // Accepting it wrote nothing while cell() went on reporting the value, which
  // is the edit disappearing with nobody told.
  assert.throws(
    () => workbook.sheets[0]?.set('A1', 'written'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.equal(workbook.sheets[0]?.cell('A1'), undefined)
  assert.deepEqual(workbook.sheets[0]?.mergedRanges, []) // no bytes to read merges from
  assert.equal(workbook.sheets[0]?.rowHeight(1), undefined)
  assert.equal(workbook.sheets[0]?.columnWidth('A'), undefined)
  assert.equal(workbook.sheets[0]?.gridlinesVisible, true) // default, no bytes to read
  assert.equal(workbook.sheets[0]?.zoomPercent, undefined)
  assert.equal(workbook.sheets[0]?.isRowHidden(1), false)
  assert.equal(workbook.sheets[0]?.isColumnHidden('A'), false)
  assert.equal(workbook.sheets[0]?.rowGroupLevel(1), 0)
  assert.equal(workbook.sheets[0]?.columnGroupLevel('A'), 0)
  assert.deepEqual(workbook.sheets[0]?.validations, [])
  assert.deepEqual(workbook.sheets[0]?.conditionalFormats, [])
  assert.throws(
    () => workbook.sheets[0]?.protect(),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.merge('A1:B2'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.setRowHeight(1, 20),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.setColumnWidth('A', 20),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.autoFilter('A1:B2'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.freeze('B2'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.hideRow(1),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.hideColumn('A'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.insertRows(1),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.insertColumns('A'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.deleteRows(1),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  assert.throws(
    () => workbook.sheets[0]?.deleteColumns('A'),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
  const gone = (error: unknown) => error instanceof XlsxError && error.code === 'missing-part'
  assert.throws(() => workbook.sheets[0]?.tabColor('FF0000'), gone)
  assert.throws(() => workbook.sheets[0]?.showGridlines(false), gone)
  assert.throws(() => workbook.sheets[0]?.showHeadings(false), gone)
  assert.throws(() => workbook.sheets[0]?.zoom(80), gone)
  assert.throws(() => workbook.sheets[0]?.groupRows(1, 2), gone)
  assert.throws(() => workbook.sheets[0]?.groupColumns('A', 'B'), gone)
  assert.throws(() => workbook.sheets[0]?.validate('A1', { list: ['x'] }), gone)
  assert.throws(
    () =>
      workbook.sheets[0]?.conditionalFormat('A1', { colorScale: { min: 'FFFFFF', max: '000000' } }),
    gone,
  )
})

test('a reference the sheet cannot hold does not break lookups around it', () => {
  // Reading is deliberately lenient about references real files get wrong, but
  // one bad cell used to make every cell() call in the sheet throw while
  // cells() returned the whole sheet happily.
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c><c r="XFE1"><v>2</v></c></row>'),
  )
  const sheet = workbook.sheets[0]

  assert.deepEqual(sheet?.cell('A1')?.value, { kind: 'number', value: 1 })
  assert.equal([...(sheet?.cells() ?? [])].length, 2)
  // Nothing legal can address it, so there is nothing to return.
  assert.equal(sheet?.cell('XFE1'), undefined)
})

test('refuses a number format when the package has no style table', () => {
  const bytes = writeContainer({
    parts: new Map([
      ['_rels/.rels', encode(ROOT_RELS)],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="D" r:id="rId1"/></sheets></workbook>'),
      ],
      ['xl/_rels/workbook.xml.rels', encode(WORKBOOK_RELS)],
      [
        'xl/worksheets/sheet1.xml',
        encode(
          '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        ),
      ],
    ]),
  })
  const workbook = readWorkbook(bytes)

  // It used to accept the write and drop the format without a word, so the
  // caller had no way to learn the cell would not display as they asked.
  assert.throws(
    () => workbook.sheets[0]?.set('A1', 0.25, { numberFormat: '0.0%' }),
    /style table|styles/i,
  )
  assert.deepEqual(workbook.sheets[0]?.cell('A1')?.value, { kind: 'number', value: 1 })
})

test('adds calcPr before the elements the schema puts after it', () => {
  // CT_Workbook is a sequence, so calcPr appended at the end sits after
  // pivotCaches and extLst, and the workbook part no longer validates.
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/workbook.xml':
          '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '<pivotCaches><pivotCache cacheId="1" r:id="rId9"/></pivotCaches>' +
          '<extLst><ext uri="{XYZ}"/></extLst></workbook>',
      },
    }),
  )
  workbook.sheets[0]?.set('A1', { formula: 'B1+1' })

  const parts = readContainer(workbook.toBytes()).parts
  const book = new TextDecoder().decode(parts.get('xl/workbook.xml') ?? new Uint8Array())

  assert.ok(
    book.indexOf('<calcPr') < book.indexOf('<pivotCaches'),
    `calcPr must precede pivotCaches: ${book}`,
  )
})

test('reports which epoch the workbook counts serials from', () => {
  const plain = readWorkbook(build('<row r="1"/>'))
  const alternate = readWorkbook(build('<row r="1"/>', { date1904: true }))

  assert.equal(plain.epoch, 1900)
  assert.equal(alternate.epoch, 1904)
})

test('reports the sheetId the workbook part gives a sheet', () => {
  const workbook = readWorkbook(build('<row r="1"/>'))

  assert.equal(workbook.sheets[0]?.sheetId, '1')
})

test('reports a cell reference canonically however the file spelled it', () => {
  // The file's own spelling made cell.reference disagree with what set() and
  // formatReference produce, so a caller keying a map on it missed.
  const workbook = readWorkbook(
    build('<row r="1"><c r="$a$1"><v>1</v></c><c r="B1"><v>2</v></c></row>'),
  )
  const cells = [...(workbook.sheets[0]?.cells() ?? [])]

  assert.deepEqual(
    cells.map((cell) => cell.reference),
    ['A1', 'B1'],
  )
  assert.equal(workbook.sheets[0]?.cell('A1')?.reference, 'A1')
})

test('a shared formula dependent says it shares rather than reporting no expression', () => {
  // It used to report formula: '', which is falsy, so `if (cell.formula)`
  // treated a dependent as a cell holding a literal.
  const workbook = readWorkbook(
    build(
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>2</v></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>',
    ),
  )
  const sheet = workbook.sheets[0]

  assert.deepEqual(sheet?.cell('A1')?.formula, { kind: 'expression', expression: 'B1*2' })
  assert.deepEqual(sheet?.cell('A2')?.formula, { kind: 'shared', master: 'A1' })
})

test('a dependent whose master is nowhere in the sheet still says it shares', () => {
  const workbook = readWorkbook(
    build('<row r="2"><c r="A2"><f t="shared" si="7"/><v>4</v></c></row>'),
  )

  assert.deepEqual(workbook.sheets[0]?.cell('A2')?.formula, { kind: 'shared' })
})

test('a refused write names the sheet it was aimed at', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(
    () => workbook.sheets[0]?.set('A1', Number.NaN),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.reference === 'A1' &&
      error.sheet === 'Data' &&
      error.part === 'xl/worksheets/sheet1.xml',
  )
})

test('a refused date write names the sheet too', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.throws(
    () => workbook.sheets[0]?.set('A1', new Date(1800, 0, 1)),
    (error: unknown) => error instanceof XlsxError && error.sheet === 'Data',
  )
})

test('an unreadable cell value names the sheet and part it is in', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>banana</v></c></row>'))

  assert.throws(
    () => [...(workbook.sheets[0]?.cells() ?? [])],
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'invalid-content' &&
      error.reference === 'A1' &&
      error.sheet === 'Data' &&
      error.part === 'xl/worksheets/sheet1.xml',
  )
})

test('writing into a merged cell that is not the anchor is refused', () => {
  const merged =
    '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
    // The second ref has no colon, which is not a range and is ignored.
    '<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="D4"/></mergeCells></worksheet>'
  const workbook = readWorkbook(build('', { extra: { 'xl/worksheets/sheet1.xml': merged } }))
  const sheet = workbook.sheets[0]

  assert.throws(
    () => sheet?.set('B2', 5),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unwritable-value' &&
      error.reference === 'B2' &&
      error.sheet === 'Data' &&
      /A1/.test(error.message),
  )
  // The anchor is the one member a value shows in, so it stays writable.
  sheet?.set('A1', 9)
  assert.equal(sheet?.cell('A1')?.value.kind, 'number')
  // A cell outside every merge is free, and so is the ignored colon-less one.
  sheet?.set('Z9', 1)
  sheet?.set('D4', 2)
  assert.equal(sheet?.cell('Z9')?.value.kind, 'number')
})

const TABLE_ROWS =
  '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
  '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>'

function tableBook(
  sheetData: string,
  tableRef = 'A1:B2',
  table = `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="${tableRef}" totalsRowShown="0"><autoFilter ref="${tableRef}"/><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns><tableStyleInfo name="TableStyleMedium9"/></table>`,
) {
  return build('', {
    extra: {
      'xl/worksheets/sheet1.xml':
        `<worksheet><dimension ref="${tableRef}"/><sheetData>${sheetData}</sheetData>` +
        '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels':
        '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
      'xl/tables/table1.xml': table,
    },
  })
}

const tablePartOf = (bytes: Uint8Array) =>
  new TextDecoder().decode(
    readContainer(bytes).parts.get('xl/tables/table1.xml') ?? new Uint8Array(),
  )

test('a write directly below a table grows the table to include it', () => {
  const workbook = readWorkbook(tableBook(TABLE_ROWS))
  workbook.sheets[0]?.set('A3', 5)

  const table = tablePartOf(workbook.toBytes())
  assert.match(table, /<table[^>]*\sref="A1:B3"/)
  assert.match(table, /<autoFilter ref="A1:B3"\/>/)
})

test('contiguous writes below a table grow it to the furthest one', () => {
  const workbook = readWorkbook(tableBook(TABLE_ROWS))
  workbook.sheets[0]?.set('B3', 5)
  workbook.sheets[0]?.set('A4', 6)

  assert.match(tablePartOf(workbook.toBytes()), /<table[^>]*\sref="A1:B4"/)
})

test('a write past a gap below a table does not grow it', () => {
  const workbook = readWorkbook(tableBook(TABLE_ROWS))
  workbook.sheets[0]?.set('A5', 5)

  assert.match(tablePartOf(workbook.toBytes()), /<table[^>]*\sref="A1:B2"/)
})

test('a write just right of a table grows it and adds a column', () => {
  const workbook = readWorkbook(tableBook(TABLE_ROWS))
  workbook.sheets[0]?.set('C1', 9)

  const table = tablePartOf(workbook.toBytes())
  assert.match(table, /<table[^>]*\sref="A1:C2"/)
  assert.match(table, /<tableColumns count="3">/)
})

test('an untouched part is copied through still compressed after an edit', async () => {
  const { readZip } = await import('./zip.js')
  const bytes = new Uint8Array(await readFile('fixtures/real/WithChart.xlsx'))
  const before = new Map(readZip(bytes).map((entry) => [entry.name, entry.compressed]))

  const workbook = readWorkbook(bytes)
  workbook.sheets[0]?.set('A100', 'edited far from the chart')
  const after = new Map(readZip(workbook.toBytes()).map((entry) => [entry.name, entry.compressed]))

  // The theme is never touched, so its bytes must be the file's own, not ours
  // re-deflated. Excel's DEFLATE differs from fflate's, so equality means the
  // compressed bytes were passed through rather than inflated and rebuilt.
  const theme = 'xl/theme/theme1.xml'
  assert.deepEqual([...(after.get(theme) ?? [])], [...(before.get(theme) ?? [])])
})
const expressionOf = (cell: Cell | undefined): string | undefined =>
  cell?.formula?.kind === 'expression' ? cell.formula.expression : undefined

test('insertRows pushes rows down and moves formulas with them', () => {
  const workbook = readWorkbook(
    build(
      '<row r="1"><c r="A1"><v>1</v></c></row>' +
        '<row r="2"><c r="A2"><v>2</v></c></row>' +
        '<row r="3"><c r="A3"><f>A1+A2</f><v>3</v></c><c r="B3"><f>A3*2</f><v>6</v></c></row>',
    ),
  )
  workbook.sheets[0]?.insertRows(2)

  const back = readWorkbook(workbook.toBytes())
  const cells = [...(back.sheets[0]?.cells() ?? [])]
  const at = (reference: string) => cells.find((cell) => cell.reference === reference)
  assert.deepEqual(at('A1')?.value, { kind: 'number', value: 1 })
  assert.equal(at('A2'), undefined)
  assert.deepEqual(at('A3')?.value, { kind: 'number', value: 2 })
  assert.equal(expressionOf(at('A4')), 'A1+A3')
  assert.equal(expressionOf(at('B4')), 'A4*2')
})

test('insertRows carries a cell edited this session along with the shift', () => {
  const workbook = readWorkbook(build('<row r="5"><c r="A5"><v>1</v></c></row>'))
  workbook.sheets[0]?.set('A5', 'moved')
  workbook.sheets[0]?.insertRows(3)

  const cells = [...(readWorkbook(workbook.toBytes()).sheets[0]?.cells() ?? [])]
  assert.deepEqual(cells.find((cell) => cell.reference === 'A6')?.value, {
    kind: 'text',
    value: 'moved',
  })
})

test('insertRows moves references from other sheets and defined names', () => {
  const workbook = createWorkbook('Data')
  workbook.sheets[0]?.set('A5', 10)
  workbook.addSheet('Calc').set('A1', { formula: 'Data!A5*2' })
  workbook.defineName('Target', 'Data!$A$5')
  workbook.sheets[0]?.insertRows(3)

  const back = readWorkbook(workbook.toBytes())
  const data = [...(back.sheet('Data')?.cells() ?? [])]
  const calc = [...(back.sheet('Calc')?.cells() ?? [])]
  assert.deepEqual(data.find((cell) => cell.reference === 'A6')?.value, {
    kind: 'number',
    value: 10,
  })
  assert.equal(expressionOf(calc.find((cell) => cell.reference === 'A1')), 'Data!A6*2')
  assert.equal(back.definedNames.get('Target'), 'Data!$A$6')
})

test('insertRows refuses a bad row, a bad count and an overflow', () => {
  const workbook = readWorkbook(build('<row r="1048576"><c r="A1048576"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  assert.throws(
    () => sheet?.insertRows(0),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => sheet?.insertRows(1, 0),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  assert.throws(
    () => sheet?.insertRows(1),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('insertRows refuses a sheet that carries a table', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        'xl/tables/table1.xml': '<table ref="A1:B2"/>',
      },
    }),
  )
  assert.throws(
    () => workbook.sheets[0]?.insertRows(1),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unsupported-edit' &&
      error.message.includes('a table'),
  )
})

test('insertRows refuses a sheet that carries a drawing', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
      },
    }),
  )
  assert.throws(
    () => workbook.sheets[0]?.insertRows(1),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unsupported-edit' &&
      error.message.includes('drawing'),
  )
})

test('insertColumns refuses a sheet that carries a pivot table', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>',
      },
    }),
  )
  assert.throws(
    () => workbook.sheets[0]?.insertColumns('A'),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unsupported-edit' &&
      error.message.includes('pivot table'),
  )
})

test('deleteColumns refuses a sheet that carries a comment', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>',
      },
    }),
  )
  assert.throws(
    () => workbook.sheets[0]?.deleteColumns('A'),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unsupported-edit' &&
      error.message.includes('comment'),
  )
})

test('insertRows allows a sheet whose relationships pin no cells', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/></Relationships>',
      },
    }),
  )
  workbook.sheets[0]?.insertRows(1)
  const cells = [...(readWorkbook(workbook.toBytes()).sheets[0]?.cells() ?? [])]
  assert.deepEqual(cells.find((cell) => cell.reference === 'A2')?.value, {
    kind: 'number',
    value: 1,
  })
})

test('insertColumns pushes columns over and moves formulas with them', () => {
  const workbook = createWorkbook('Data')
  workbook.sheets[0]?.set('A1', 1)
  workbook.sheets[0]?.set('C1', { formula: 'A1+B1' })
  workbook.addSheet('Calc').set('A1', { formula: 'Data!C1*2' })
  workbook.defineName('Target', 'Data!$C$1')
  workbook.sheets[0]?.insertColumns('B')

  const back = readWorkbook(workbook.toBytes())
  const data = [...(back.sheet('Data')?.cells() ?? [])]
  const at = (reference: string) => data.find((cell) => cell.reference === reference)
  assert.deepEqual(at('A1')?.value, { kind: 'number', value: 1 })
  assert.equal(at('C1'), undefined)
  assert.equal(expressionOf(at('D1')), 'A1+C1')
  const calc = [...(back.sheet('Calc')?.cells() ?? [])]
  assert.equal(expressionOf(calc.find((cell) => cell.reference === 'A1')), 'Data!D1*2')
  assert.equal(back.definedNames.get('Target'), 'Data!$D$1')
})

test('insertColumns takes a cols width entry and an implicit cell in its stride', () => {
  const workbook = readWorkbook(
    build(
      '<cols><col min="1" max="1" width="8"/></cols>' +
        '<row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c></row>',
    ),
  )
  workbook.sheets[0]?.insertColumns('A')
  const cells = [...(readWorkbook(workbook.toBytes()).sheets[0]?.cells() ?? [])]
  assert.deepEqual(cells.find((cell) => cell.reference === 'B1')?.value, {
    kind: 'number',
    value: 1,
  })
})

test('insertColumns refuses a bad count, an off-sheet column and a table', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="XFD1"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  assert.throws(
    () => sheet?.insertColumns('A', 0),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  assert.throws(
    () => sheet?.insertColumns('A'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  assert.throws(
    () => sheet?.insertColumns('XFE'),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  const tabled = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        'xl/tables/table1.xml': '<table ref="A1:B2"/>',
      },
    }),
  )
  assert.throws(
    () => tabled.sheets[0]?.insertColumns('A'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unsupported-edit',
  )
})

test('deleteRows pulls rows up, #REF!s a lost cell and shrinks an overlapping range', () => {
  const workbook = createWorkbook('Data')
  const sheet = workbook.sheets[0]
  sheet?.set('A1', 1)
  sheet?.set('A2', 2)
  sheet?.set('A4', { formula: 'A2+SUM(A1:A4)' })
  workbook.addSheet('Calc').set('A1', { formula: 'Data!A2+Data!A4' })
  sheet?.deleteRows(2)

  const back = readWorkbook(workbook.toBytes())
  const data = [...(back.sheet('Data')?.cells() ?? [])]
  const at = (reference: string) => data.find((cell) => cell.reference === reference)
  assert.deepEqual(at('A1')?.value, { kind: 'number', value: 1 })
  assert.equal(at('A2'), undefined)
  assert.equal(expressionOf(at('A3')), '#REF!+SUM(A1:A3)')
  const calc = [...(back.sheet('Calc')?.cells() ?? [])]
  assert.equal(expressionOf(calc.find((cell) => cell.reference === 'A1')), '#REF!+Data!A3')
})

test('deleteRows refuses a bad row, a bad count and a collapsing merge', () => {
  const workbook = readWorkbook(
    build(
      '<row r="5"><c r="A5"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells><sheetData>',
    ),
  )
  const sheet = workbook.sheets[0]
  assert.throws(
    () => sheet?.deleteRows(0),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => sheet?.deleteRows(1, 0),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  assert.throws(
    () => sheet?.deleteRows(5),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('deleteRows refuses destroying a shared formula master and a sheet with a table', () => {
  const master = readWorkbook(
    build('<row r="2"><c r="B2"><f t="shared" ref="B2:B9" si="0">A2</f></c></row>'),
  )
  assert.throws(
    () => master.sheets[0]?.deleteRows(2),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  const tabled = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        'xl/tables/table1.xml': '<table ref="A1:B2"/>',
      },
    }),
  )
  assert.throws(
    () => tabled.sheets[0]?.deleteRows(1),
    (error: unknown) => error instanceof XlsxError && error.code === 'unsupported-edit',
  )
})

test('deleteRows keeps a merge it only clips and a master above the deletion', () => {
  const workbook = readWorkbook(
    build(
      '<row r="1"><c r="B1"><f t="shared" ref="B1:B9" si="0">A1</f></c></row>' +
        '<row r="6"><c r="A6"><v>1</v></c></row>' +
        '</sheetData><mergeCells count="1"><mergeCell ref="A4:A6"/></mergeCells><sheetData>',
    ),
  )
  workbook.sheets[0]?.deleteRows(5)
  const out = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.match(out, /<mergeCell ref="A4:A5"\/>/)
  assert.match(out, /<row r="5"><c r="A5"><v>1<\/v>/)
})

test('deleteColumns pulls columns in, #REF!s a lost cell and shrinks an overlapping range', () => {
  const workbook = createWorkbook('Data')
  const sheet = workbook.sheets[0]
  sheet?.set('A1', 1)
  sheet?.set('B1', 2)
  sheet?.set('D1', { formula: 'B1+SUM(A1:D1)' })
  workbook.addSheet('Calc').set('A1', { formula: 'Data!B1+Data!D1' })
  sheet?.deleteColumns('B')

  const back = readWorkbook(workbook.toBytes())
  const data = [...(back.sheet('Data')?.cells() ?? [])]
  const at = (reference: string) => data.find((cell) => cell.reference === reference)
  assert.deepEqual(at('A1')?.value, { kind: 'number', value: 1 })
  assert.equal(at('B1'), undefined)
  assert.equal(expressionOf(at('C1')), '#REF!+SUM(A1:C1)')
  const calc = [...(back.sheet('Calc')?.cells() ?? [])]
  assert.equal(expressionOf(calc.find((cell) => cell.reference === 'A1')), '#REF!+Data!C1')
})

test('deleteColumns refuses a bad column, a bad count and a collapsing master', () => {
  const master = readWorkbook(
    build('<row r="1"><c r="B1"><f t="shared" ref="B1:B9" si="0">A1</f></c></row>'),
  )
  assert.throws(
    () => master.sheets[0]?.deleteColumns('XFE'),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => master.sheets[0]?.deleteColumns('A', 0),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
  assert.throws(
    () => master.sheets[0]?.deleteColumns('B'),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('link writes an internal link inline and an external link through a relationship', () => {
  const workbook = createWorkbook('Data')
  workbook.sheets[0]?.set('A1', 'home')
  workbook.sheets[0]?.set('A2', 'site')
  workbook.addSheet('Other')
  workbook.sheets[0]?.link('A1', { location: 'Other!B2', tooltip: 'jump' })
  workbook.sheets[0]?.link('A2', { url: 'https://example.com/a?x=1&y=2' })

  const parts = readContainer(workbook.toBytes()).parts
  const sheet = decode(parts.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<hyperlink ref="A1" location="Other!B2" tooltip="jump"\/>/)
  assert.match(sheet, /<hyperlink ref="A2" r:id="rId1"\/>/)
  assert.match(sheet, /<worksheet [^>]*xmlns:r=/)
  const rels = decode(parts.get('xl/worksheets/_rels/sheet1.xml.rels'))
  assert.match(
    rels,
    /Id="rId1" Type="[^"]*hyperlink" Target="https:\/\/example.com\/a\?x=1&amp;y=2" TargetMode="External"/,
  )
})

test('link replaces a link on the same cell and moves with an inserted row', () => {
  const workbook = readWorkbook(build('<row r="3"><c r="A3"><v>1</v></c></row>'))
  const sheet = workbook.sheets[0]
  sheet?.link('A3', { location: 'old' })
  sheet?.link('A3', { location: 'new' })
  sheet?.insertRows(1)

  const out = decode(readContainer(workbook.toBytes()).parts.get('xl/worksheets/sheet1.xml'))
  assert.doesNotMatch(out, /location="old"/)
  assert.match(out, /<hyperlink ref="A4" location="new"\/>/)
})

test('link refuses a bad reference, an empty target and a missing sheet part', () => {
  const workbook = createWorkbook()
  const sheet = workbook.sheets[0]
  assert.throws(
    () => sheet?.link('nope', { url: 'https://example.com' }),
    (error: unknown) => error instanceof XlsxError && error.code === 'bad-reference',
  )
  assert.throws(
    () => sheet?.link('A1', { url: '' }),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('link appends its relationship after ones the sheet already has', () => {
  const workbook = readWorkbook(
    build('<row r="1"><c r="A1"><v>1</v></c></row>', {
      extra: {
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
      },
    }),
  )
  workbook.sheets[0]?.link('A1', { url: 'https://example.com' })
  const parts = readContainer(workbook.toBytes()).parts
  const rels = decode(parts.get('xl/worksheets/_rels/sheet1.xml.rels'))
  assert.match(rels, /Id="rId1" Type="[^"]*drawing"/)
  assert.match(rels, /Id="rId2" Type="[^"]*hyperlink" Target="https:\/\/example.com"/)
  assert.match(decode(parts.get('xl/worksheets/sheet1.xml')), /<hyperlink ref="A1" r:id="rId2"\/>/)
})
