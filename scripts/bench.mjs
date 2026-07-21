// Writes a lot of cells and interleaves reads, which is the shape of work that
// turns a quadratic path into a hang. Run it before and after touching the
// write path.
import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const { register } = require('tsx/esm/api')
const unregister = register()

const { readWorkbook } = await import('../src/lib/document.ts')
const { writeContainer } = await import('../src/lib/container.ts')

const encode = (text) => new TextEncoder().encode(text)

function workbookOf(rows) {
  const body = Array.from(
    { length: rows },
    (_unused, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`,
  ).join('')

  return writeContainer({
    parts: new Map([
      [
        '_rels/.rels',
        encode(
          '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        ),
      ],
      [
        'xl/workbook.xml',
        encode('<workbook><sheets><sheet name="Data" r:id="rId1"/></sheets></workbook>'),
      ],
      [
        'xl/_rels/workbook.xml.rels',
        encode(
          '<Relationships>' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
            '</Relationships>',
        ),
      ],
      ['xl/worksheets/sheet1.xml', encode(`<worksheet><sheetData>${body}</sheetData></worksheet>`)],
      [
        'xl/styles.xml',
        encode(
          '<styleSheet><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>',
        ),
      ],
    ]),
  })
}

const time = (label, run) => {
  const started = process.hrtime.bigint()
  run()
  const ms = Number(process.hrtime.bigint() - started) / 1e6
  console.log(`${label.padEnd(42)} ${ms.toFixed(0).padStart(7)} ms`)
  return ms
}

for (const size of [2500, 5000, 10000]) {
  const bytes = workbookOf(size)
  console.log(`\n${size} rows`)

  time(`  write ${size} cells`, () => {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    for (let row = 1; row <= size; row++) sheet.set(`B${row}`, row)
    workbook.toBytes()
  })

  time(`  write ${size} cells, reading between each`, () => {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    for (let row = 1; row <= size; row++) {
      sheet.set(`B${row}`, row)
      sheet.cell(`A${row}`)
    }
    workbook.toBytes()
  })

  time(`  write ${size} dates`, () => {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    for (let row = 1; row <= size; row++) sheet.set(`B${row}`, new Date(2024, 0, 1))
    workbook.toBytes()
  })

  time(`  append ${size} new rows`, () => {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    for (let row = 1; row <= size; row++) sheet.set(`A${size + row}`, row)
    workbook.toBytes()
  })
}

await unregister()
