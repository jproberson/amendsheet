import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer, writeContainer } from './container.js'
import { readWorkbook } from './document.js'

const encode = (text: string) => new TextEncoder().encode(text)

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
  assert.equal(
    cell?.value.kind === 'date' && cell.value.value.toISOString(),
    '2024-01-01T00:00:00.000Z',
  )
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

  assert.equal(
    cell?.value.kind === 'date' && cell.value.value.toISOString(),
    '1904-01-01T00:00:00.000Z',
  )
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

test('round trips every corpus file without losing a part', async () => {
  const files = (await readdir('corpus/real')).filter((name) => name.endsWith('.xlsx'))
  const damaged: string[] = []

  for (const file of files) {
    const original = new Uint8Array(await readFile(`corpus/real/${file}`))
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

test('reads cells from every corpus file', async () => {
  const files = (await readdir('corpus/real')).filter((name) => name.endsWith('.xlsx'))
  let cells = 0
  let dates = 0

  for (const file of files) {
    const workbook = readWorkbook(new Uint8Array(await readFile(`corpus/real/${file}`)))
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
  workbook.sheets[0]?.set('A1', new Date('2024-01-01T00:00:00Z'))

  const reopened = readWorkbook(workbook.toBytes())
  const [cell] = [...(reopened.sheets[0]?.cells() ?? [])]

  assert.equal(cell?.value.kind, 'date')
  assert.equal(
    cell?.value.kind === 'date' && cell.value.value.toISOString(),
    '2024-01-01T00:00:00.000Z',
  )
})
