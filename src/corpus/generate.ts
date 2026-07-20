import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { toArrayBuffer } from '../harness/zip.js'
import { QUIRKS } from './quirks.js'

const CORPUS_DIR = join(process.cwd(), 'corpus')

/** 1x1 transparent PNG, so the corpus needs no binary assets checked in. */
const TRANSPARENT_PIXEL_PNG = toArrayBuffer(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
    'base64',
  ),
)

/**
 * A workbook using the features a real business template depends on. This is
 * the file that answers the central question: if you open it, change nothing,
 * and save it, what survives?
 */
async function richWorkbook(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'corpus-harness'
  workbook.created = new Date('2024-01-01T00:00:00Z')

  const worksheet = workbook.addWorksheet('Report', {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }],
  })

  worksheet.mergeCells('A1:C1')
  const title = worksheet.getCell('A1')
  title.value = 'Quarterly Report'
  title.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } }
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } }
  title.alignment = { horizontal: 'center', vertical: 'middle' }

  worksheet.getRow(2).values = ['Region', 'Revenue', 'Approved']
  worksheet.getRow(2).font = { bold: true }
  worksheet.getRow(2).border = { bottom: { style: 'medium' } }

  const rows: Array<[string, number, string]> = [
    ['North', 12500.5, 'yes'],
    ['South', 8300, 'no'],
    ['East', 19750.25, 'yes'],
  ]
  rows.forEach((r, i) => {
    const row = worksheet.getRow(3 + i)
    row.values = r
    row.getCell(2).numFmt = '"$"#,##0.00'
  })

  const total = worksheet.getCell('B6')
  total.value = { formula: 'SUM(B3:B5)', result: 40550.75 }
  total.numFmt = '"$"#,##0.00'
  total.font = { bold: true }
  worksheet.getCell('A6').value = 'Total'

  worksheet.getCell('A8').value = new Date('2024-03-31T00:00:00Z')
  worksheet.getCell('A8').numFmt = 'yyyy-mm-dd'

  worksheet.getColumn(1).width = 18
  worksheet.getColumn(2).width = 14

  worksheet.autoFilter = 'A2:C2'

  worksheet.addConditionalFormatting({
    ref: 'B3:B5',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['10000'],
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6EFCE' } } },
      },
    ],
  })

  for (let r = 3; r <= 5; r++) {
    worksheet.getCell(`C${r}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"yes,no"'],
    }
  }

  worksheet.getCell('A3').note = 'Includes cross-border revenue'

  const imageId = workbook.addImage({ buffer: TRANSPARENT_PIXEL_PNG, extension: 'png' })
  worksheet.addImage(imageId, 'E2:F6')

  workbook.definedNames.add('Report!$B$3:$B$5', 'RevenueRange')

  const lookups = workbook.addWorksheet('Lookups')
  lookups.getCell('A1').value = 'yes'
  lookups.getCell('A2').value = 'no'

  return new Uint8Array(await workbook.xlsx.writeBuffer())
}

/** Plain values only: the control case. If this loses data, something is very wrong. */
async function simpleWorkbook(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Data')
  worksheet.addRow(['id', 'name', 'score'])
  for (let i = 1; i <= 25; i++) worksheet.addRow([i, `name-${i}`, i * 1.5])
  return new Uint8Array(await workbook.xlsx.writeBuffer())
}

async function main() {
  await mkdir(join(CORPUS_DIR, 'generated'), { recursive: true })
  await mkdir(join(CORPUS_DIR, 'quirks'), { recursive: true })
  await mkdir(join(CORPUS_DIR, 'manual'), { recursive: true })

  const written: string[] = []

  const rich = await richWorkbook()
  await writeFile(join(CORPUS_DIR, 'generated', 'rich-exceljs.xlsx'), rich)
  written.push('generated/rich-exceljs.xlsx')

  const simple = await simpleWorkbook()
  await writeFile(join(CORPUS_DIR, 'generated', 'simple-exceljs.xlsx'), simple)
  written.push('generated/simple-exceljs.xlsx')

  for (const quirk of QUIRKS) {
    const file = join(CORPUS_DIR, 'quirks', `${quirk.name}.xlsx`)
    await writeFile(file, quirk.build())
    written.push(`quirks/${quirk.name}.xlsx`)
  }

  await writeFile(
    join(CORPUS_DIR, 'manual', 'README.md'),
    [
      '# Manual corpus',
      '',
      'Drop real `.xlsx` files here. Diversity of *producer* matters far more than count.',
      '',
      'Aim for at least one of each:',
      '',
      '- Excel (desktop, current) — save a file with a chart and a pivot table',
      '- Excel (older, .xls saved as .xlsx)',
      '- LibreOffice Calc — export to xlsx',
      '- Google Sheets — File > Download > Microsoft Excel',
      '- Apple Numbers — Export to Excel',
      '- A server-generated report (openpyxl, Apache POI, a BI tool export)',
      '- Government open data — search any open data portal for "xlsx"',
      '',
      'Files here are gitignored by default: they may be large or non-redistributable.',
      'Record provenance in a sibling `.txt` if you plan to publish the corpus.',
    ].join('\n'),
  )

  console.log(`Wrote ${written.length} corpus files to ${CORPUS_DIR}`)
  for (const file of written) console.log(`  ${file}`)
  console.log('\nAdd real-world files to corpus/manual/ — see its README.')
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
