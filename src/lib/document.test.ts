import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer, writeContainer } from './container.js'
import { readWorkbook } from './document.js'
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
  options: { date1904?: boolean; extra?: Record<string, string> } = {},
) {
  const parts: Record<string, Uint8Array> = {
    '_rels/.rels': encode(ROOT_RELS),
    'xl/workbook.xml': encode(
      `<workbook><workbookPr${options.date1904 ? ' date1904="1"' : ''}/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': encode(WORKBOOK_RELS),
    'xl/styles.xml': encode(STYLES),
    'xl/worksheets/sheet1.xml': encode(
      `<worksheet><sheetData>${sheetBody}</sheetData></worksheet>`,
    ),
    'xl/charts/chart1.xml': encode('<chart/>'),
  }
  for (const [path, content] of Object.entries(options.extra ?? {})) parts[path] = encode(content)
  return writeContainer({ parts: new Map(Object.entries(parts)) })
}

test('exposes sheets by name', () => {
  const workbook = readWorkbook(build('<row r="1"><c r="A1"><v>1</v></c></row>'))

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Data'],
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
    fill: { color: 'FF0000' },
    border: { top: { style: 'thin' } },
  })

  const cell = workbook.sheets[0]?.cell('A1')
  assert.deepEqual(cell?.font, { bold: true })
  assert.deepEqual(cell?.fill, { color: 'FFFF0000' })
  assert.deepEqual(cell?.border, { top: { style: 'thin' } })
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
  workbook.sheets[0]?.set('A1', 'x', { fill: { color: '00FF00' }, font: { bold: true } })

  const styles = decode(readContainer(workbook.toBytes()).parts.get('xl/styles.xml'))
  assert.match(styles, /<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"\//)
  const xf = [...styles.matchAll(/<xf [^>]*\/>/g)].map((match) => match[0]).at(-1) ?? ''
  assert.match(xf, /applyFill="1"/)
  assert.match(xf, /applyFont="1"/)
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
