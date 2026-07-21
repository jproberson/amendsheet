import { writeParts, encodeText } from '../harness/zip.js'

/**
 * Hand-built .xlsx files exercising legal-but-unusual constructs that real
 * producers (LibreOffice, Google Sheets, server-side report generators) emit
 * and that naive parsers get wrong. Each one is a valid document; a reader
 * that fails on these is failing on real files, not synthetic ones.
 */

const CONTENT_TYPES = (extra: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${extra}</Types>`

const SHARED_STRINGS_CONTENT_TYPE = `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n`

const ROOT_RELATIONSHIPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELATIONSHIPS = (
  hasSharedStrings: boolean,
  extra = '',
) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
${hasSharedStrings ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ''}
${extra}</Relationships>`

const WORKBOOK = (date1904: boolean) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr${date1904 ? ' date1904="1"' : ''}/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

/** Style index 1 is a date format; index 0 is General. */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`

const sheet = (body: string, dimension = 'A1:C5') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="${dimension}"/>
<sheetData>${body}</sheetData>
</worksheet>`

function buildWorkbook(options: {
  sheetXml: string
  sharedStrings?: string
  date1904?: boolean
  /** Extra content type declarations, and the parts they describe. */
  contentTypes?: string
  extraRelationships?: string
  extraParts?: Record<string, Uint8Array>
}): Uint8Array {
  const { sharedStrings } = options
  const declared = (sharedStrings ? SHARED_STRINGS_CONTENT_TYPE : '') + (options.contentTypes ?? '')
  const parts: Record<string, Uint8Array> = {
    '[Content_Types].xml': encodeText(CONTENT_TYPES(declared)),
    '_rels/.rels': encodeText(ROOT_RELATIONSHIPS),
    'xl/workbook.xml': encodeText(WORKBOOK(options.date1904 ?? false)),
    'xl/_rels/workbook.xml.rels': encodeText(
      WORKBOOK_RELATIONSHIPS(sharedStrings !== undefined, options.extraRelationships ?? ''),
    ),
    'xl/styles.xml': encodeText(STYLES),
    'xl/worksheets/sheet1.xml': encodeText(options.sheetXml),
  }
  if (sharedStrings !== undefined) {
    parts['xl/sharedStrings.xml'] = encodeText(sharedStrings)
  }
  for (const [path, bytes] of Object.entries(options.extraParts ?? {})) parts[path] = bytes
  return writeParts(parts)
}

const sharedStringTable = (items: string[]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${items.length}" uniqueCount="${items.length}">
${items.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`

export const QUIRKS: Array<{
  name: string
  description: string
  /** Defaults to xlsx. A macro-bearing workbook has to be xlsm to be legal. */
  extension?: string
  build: () => Uint8Array
}> = [
  {
    name: 'macro-enabled',
    description: 'A macro project, which is a binary part the library never interprets',
    extension: 'xlsm',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(`<row r="1"><c r="A1"><v>1</v></c></row>`, 'A1:A1'),
        contentTypes:
          '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>\n' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>\n',
        extraRelationships:
          '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>\n',
        // Not a real compiled project. What is being measured is that a binary
        // part the library has no opinion about survives byte for byte.
        extraParts: {
          'xl/vbaProject.bin': new Uint8Array([
            0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x01, 0x02, 0x03, 0xfe, 0xff,
          ]),
        },
      }),
  },
  {
    name: 'inline-strings',
    description: 'Text stored as t="inlineStr" rather than in the shared string table',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c><c r="B1"><v>42</v></c></row>` +
            `<row r="2"><c r="A2" t="inlineStr"><is><t>World</t></is></c><c r="B2"><v>7.5</v></c></row>`,
          'A1:B2',
        ),
      }),
  },
  {
    name: 'missing-cell-refs',
    description: 'Rows and cells omit their r= address attributes; position is implied',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row><c t="inlineStr"><is><t>alpha</t></is></c><c><v>1</v></c></row>` +
            `<row><c t="inlineStr"><is><t>beta</t></is></c><c><v>2</v></c></row>`,
          'A1:B2',
        ),
      }),
  },
  {
    name: 'rows-out-of-order',
    description: 'Row 3 is serialized before row 1',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="3"><c r="A3"><v>300</v></c></row>` +
            `<row r="1"><c r="A1"><v>100</v></c></row>` +
            `<row r="2"><c r="A2"><v>200</v></c></row>`,
          'A1:A3',
        ),
      }),
  },
  {
    name: 'lying-dimension',
    description: 'dimension claims A1:A1 but data extends to C3',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>` +
            `<row r="3"><c r="C3"><v>9</v></c></row>`,
          'A1:A1',
        ),
      }),
  },
  {
    name: 'date-epoch-1904',
    description: 'workbookPr date1904="1"; serial 40000 means a different day than the 1900 epoch',
    build: () =>
      buildWorkbook({
        date1904: true,
        sheetXml: sheet(
          `<row r="1"><c r="A1" s="1"><v>40000</v></c><c r="B1" s="1"><v>1</v></c></row>`,
          'A1:B1',
        ),
      }),
  },
  {
    name: 'date-1900-leap-bug',
    description: 'Serial 60 is Excel’s nonexistent 1900-02-29; serial 61 is 1900-03-01',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="A1" s="1"><v>59</v></c><c r="B1" s="1"><v>60</v></c><c r="C1" s="1"><v>61</v></c></row>`,
          'A1:C1',
        ),
      }),
  },
  {
    name: 'shared-formula',
    description: 'One master <f t="shared"> with dependents referencing it by si index',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>` +
            `<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>` +
            `<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>`,
          'A1:B3',
        ),
      }),
  },
  {
    name: 'sparse-columns',
    description: 'Column AA and beyond, exercising bijective base-26 reference parsing',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="Z1"><v>26</v></c><c r="AA1"><v>27</v></c><c r="AB1"><v>28</v></c><c r="BZ1"><v>78</v></c></row>`,
          'Z1:BZ1',
        ),
      }),
  },
  {
    name: 'shared-strings-with-entities',
    description: 'Strings containing XML-significant characters and a non-BMP emoji',
    build: () =>
      buildWorkbook({
        sharedStrings: sharedStringTable([
          'a &amp; b',
          '&lt;tag&gt;',
          'quote &quot;q&quot;',
          '\u{1F600} emoji',
        ]),
        sheetXml: sheet(
          `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>`,
          'A1:D1',
        ),
      }),
  },
  {
    name: 'boolean-and-error-cells',
    description: 'Typed cells: t="b" boolean and t="e" error values',
    build: () =>
      buildWorkbook({
        sheetXml: sheet(
          `<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c><c r="C1" t="e"><v>#DIV/0!</v></c><c r="D1" t="e"><v>#N/A</v></c></row>`,
          'A1:D1',
        ),
      }),
  },
]
