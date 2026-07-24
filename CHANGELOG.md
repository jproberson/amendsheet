# Changelog

Follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and semantic
versioning. The API is pre-1.0, so it is still free to change between minor
versions.

## 0.1.0 — 2026-07-24

First published release.

### Added

- `worksheet.setRowHeight(row, height)` sizes a row in points, and
  `worksheet.setColumnWidth(column, width)` sizes a column in Excel's width units.
  A width lands in the `cols` range covering its column, splitting a range that
  spans more than the one column so the rest keeps its own width. Both refuse a
  negative size, and the row refuses a number below 1.
- `worksheet.merge('A1:B2')` merges a range, joining any mergeCells the sheet
  already has and refusing a range that is not two references either side of a
  colon. A written value in a non-anchor cell of the new merge is refused, as it
  already was for a merge the file was read with. Covered values are left as they
  are, the way Excel keeps them.
- `worksheet.protect(options?)` turns on worksheet protection — the thing that
  makes a cell's `locked` and `hidden` flags bite. With no argument it matches
  Excel's Protect Sheet default; `options` names the actions that stay permitted.
  Written as a `sheetProtection` element after `sheetData`, replacing any the
  sheet already had. No password is written. It is this library's first edit to
  something other than a cell. `worksheet.unprotect()` drops it again, and
  `worksheet.protection` reads it back, reflecting a pending change as well as the
  file.
- `set(reference, value, { protection })` and `format(reference, { protection })`
  set a cell's `locked` and `hidden` flags, which take effect once the worksheet
  is protected. Written as an `xf` child after any alignment, merged onto the
  protection the cell has, and read back as `cell.protection`.
- A colour on a font, fill or border can be a theme reference `{ theme, tint }`
  or an indexed palette entry `{ indexed }`, alongside the existing `RRGGBB` /
  `AARRGGBB` hex string, as the `Color` union. A theme or indexed colour is a
  reference, so it is read back and written as one rather than resolved to the
  hex it displays as. Reading a malformed theme or index drops it, the way an
  unreadable attribute is dropped.
- Pattern fills, as the second arm of the `FillFormat` union: `{ type: 'pattern',
  pattern, color, background }` for an ECMA-376 pattern type (`lightGrid`,
  `darkTrellis`, and the rest), `color` being the pattern's foreground. Written,
  and read back off a cell.
- A cell's `font` carries `strike`, `verticalAlign` (`superscript`/`subscript`) and a
  richer `underline` — `true`, or one of `double`, `singleAccounting`,
  `doubleAccounting` — alongside the existing bold, italic, size, colour and name.
  Each is written, merged onto the font the cell has, and read back off it.
- `set(reference, value, { alignment })` and `format(reference, { alignment })`
  place a cell's text: `horizontal`, `vertical`, `wrapText`, `textRotation`
  (0–180, or 255 to stack) and `indent`, merged onto the alignment the cell
  already has. `cell.alignment` reads it back. A rotation outside the range or a
  negative indent is refused with `unwritable-value`, naming the cell.
- A cell written directly below or to the right of a table grows the table to
  include it, the way Excel does, so the new row or column is inside the table
  rather than stranded next to it. Growing sideways adds a column, named so it
  collides with none the table already has. A table with a totals row, or one
  another table sits against, is left as it was.
- `error.sheet`, the worksheet name a failure happened on, alongside `part` and
  `reference`. Every error now carries the tightest location it knows: a refused
  write names the sheet it was aimed at, and an unreadable cell value names the
  sheet and part it sits in rather than nothing.
- `worksheet.sheetId`, as the workbook part spells it, so a sheet can be matched
  against a defined name or a part this library does not interpret.
- `invalid-content`, for a part that is well formed xml but says something no
  reader can honour. `XlsxErrorCode` is now an open union: new codes arrive in
  minor releases and a switch over it needs a default. Which codes mean the
  caller is at fault and which mean the file is is written down in the README.
- `part-too-large`, for a part that decompresses to more bytes than the runtime
  can hold in one buffer, reported against the part instead of surfacing as a
  bare allocation error.
- `set(reference, { formula })` writes a formula. Written without a cached
  result, so the cell reads back empty until a spreadsheet application opens the
  file and calculates it. The workbook is marked for recalculation.
- `set(reference, value, { numberFormat })` chooses a number format code,
  reusing a format the file already has where one matches.
- `npm run external` converts every real test file with LibreOffice, before and
  after an edit, to check the output against an implementation that is not this
  one.
- `npm run browser` runs the built library through a read, an edit and a write
  inside a headless Chrome, so the browser-support claim is checked by execution
  rather than by a grep for Node APIs. It drives Chrome over the DevTools
  protocol, so it adds no dependency, and `verify.sh` runs it.
- `npm run validate` edits every fixture and checks the output against the
  integrity rules a spreadsheet application repairs a file for — a dangling
  relationship, a part with no content type, a table whose shape is
  inconsistent. It reports only a problem an edit introduces, and `verify.sh`
  runs it.
- The round-trip harness measures an edited file as well as an untouched one,
  and reports parts rewritten separately from data lost.

### Changed

- A solid fill is now `{ type: 'solid', color }` rather than `{ color }`, so the
  `FillFormat` union can name a pattern as well. `cell.fill` reads back the same
  shape.
- A part the edit does not touch is copied through still compressed, never
  inflated or rebuilt. Editing one cell in a workbook full of images or other
  sheets no longer pays to decompress and recompress them: over ~86MB of
  untouched parts, a one-cell edit went from 2.2s to 3ms and about 90MB less
  peak memory. The ZIP layer is now ours; fflate only deflates and inflates the
  parts that change.
- ZIP64 is read and written, so a workbook with more than 65535 parts, or a part
  or archive past 4GB, is handled rather than refused. A package within the plain
  ZIP limits — every ordinary workbook — comes out byte for byte as before; the
  ZIP64 records only appear when a count, size or offset needs them.
- A worksheet is read, edited and written through its raw bytes, never decoded
  into one string. A sheet whose xml is larger than V8's ~512MB string limit —
  a few million rows — used to die on a bare `RangeError` before parsing; now the
  ceiling is the largest single buffer the runtime allocates, and a part past
  even that is refused with `part-too-large`.
- `cell.formula` is a union rather than a string. A cell following a shared
  formula reported `''`, which is falsy, so `if (cell.formula)` treated it as a
  cell holding a literal. It now reports `{ kind: 'shared', master }`, naming the
  cell the expression is stored on, against `{ kind: 'expression', expression }`
  for a cell that has one.
- `cell.reference` is canonical, so it always equals
  `formatReference(cell.address)`. It used to be whatever the file wrote, so a
  file spelling a cell `$A$1` or `a1` produced a key that matched nothing the
  caller could compute, and a different key per writing application.
- `workbook.date1904` is now `workbook.epoch`, which is `1900` or `1904`.
- `WriteOptions` is now `SetOptions`. It is `set`'s options, not `toBytes`'s.
- `worksheet.cells()` returns `Iterable<Cell>` rather than `IterableIterator`.
- `columnToIndex` refuses anything that is not one or more letters, rather than
  returning `0` for `''` and a meaningless number for `'a1'`. It still converts
  columns past the last one a sheet can hold, because reading is lenient and
  that conversion goes through here.
- Reading a cell whose numeric text is not a number reports `invalid-content`
  rather than `unwritable-value`. Nobody was trying to write anything.
- Cell styles are resolved when `set` is called rather than when the file is
  written, so a read of an edited workbook agrees with a read of the bytes it
  produces. Writing a date, or a number with a date format, previously read back
  as a plain number until the file was saved and reopened.
- Bulk writes are linear in the number of edits. Interleaving reads and writes
  over 10,000 cells went from 7m46s to 69ms.
- Chartsheets and dialogsheets are no longer listed in `workbook.sheets`. They
  hold no cells, so they arrived as empty worksheets that refused writes. Their
  parts are still preserved.
- `set` refuses an edit the format cannot hold at the call that caused it,
  rather than at `toBytes()`. `NaN`, an infinity, a character XML cannot encode,
  a date outside the workbook's epoch and overwriting the master of a shared
  formula previously left a workbook that threw on every later read and could
  not be saved at all. A refused edit is not recorded, so the rest of the batch
  survives.
- Refusals name the cell they came from. A date the epoch cannot hold reported
  only the date.

### Fixed

- A `t="s"` cell pointing at a shared-string index the table does not hold is
  reported as located invalid content instead of read as empty text. A negative,
  fractional, non-numeric or past-the-end index silently became an empty string,
  losing the corruption rather than naming it.
- A `t="d"` ISO literal whose time field is out of range is left as text instead
  of silently shifting the day. An hour past 23 rolled into the next day and a
  fraction that rounded up to a full second into the next minute, and only the
  year and month were checked, so the wrong day was read as a real date.
- An archive of exactly 65535 entries written without a ZIP64 record now reads.
  0xffff is a legal entry count that fits the plain end record, so a conforming
  writer stores it directly; the reader treated the value as a ZIP64 sentinel and
  demanded a locator that was never required. A file that instead lies about its
  count is still caught while reading the entries it does not have.
- A crafted or truncated ZIP64 extra field is reported as a located not-a-zip
  instead of crashing with a bare `RangeError`. An entry that maxed a 32-bit
  size or offset but carried too few u64s behind it — or an extra length that ran
  off the end of the archive — read past the buffer, throwing outside the error
  contract with no `code` and no location.
- Editing a cell keeps a theme or indexed colour its font, fill or border already
  carried. Only an `rgb` colour was read, so the colour was dropped when the style
  was merged and rewritten — making an accent-coloured cell bold turned it
  colourless. Default text is `<color theme="1"/>`, so this fired on ordinary
  files.
- A font, fill or border added to a `styles.xml` written with a namespace prefix
  (`<x:styleSheet>`) is now prefixed throughout. Only the outer opening tag was
  rewritten, so the new element closed and nested unprefixed — `<x:font><b/></font>`
  — which is malformed and made a stricter reader reject the file.
- An OLE2 compound file — a password-protected `.xlsx` or a legacy `.xls` — is
  reported as what it is, naming decryption or conversion, rather than as a bare
  "not a zip archive" that pointed at the wrong problem.
- A UTF-16 part is refused with a clear message instead of decoding to
  interleaved-null garbage. A UTF-16 part with no byte-order mark used to read as
  an empty part with no error, since its nulls are valid UTF-8.
- Writing into a merged cell that is not the anchor is refused, the way an array
  or shared-formula anchor already was. A value stored on any other member of a
  merge never shows in any reader, so the write was silently lost.

- A `Date` written and read back is now identical to the millisecond. A serial
  is a count of days, so the conversion back was fractional and the `Date`
  constructor truncated it, leaving about one time of day in twenty a
  millisecond early.

- An attribute is rewritten whatever quotes the file used. A single-quoted
  `numFmtId` gained a second copy rather than being rewritten, and a duplicate
  attribute is fatally malformed, so `styles.xml` stopped parsing. The `count`
  on `cellXfs`, `numFmts` and `sst` had the same fault.

- `calcPr` is added where the schema expects it. `CT_Workbook` is a sequence and
  it was appended last, so a workbook with `pivotCaches` or `extLst` and no
  `calcPr` became schema invalid the first time a formula was written.

- Passing an object that is not a `Date` and names no string formula is refused
  with `unwritable-value`. It used to throw a `TypeError` from inside, outside
  the error contract entirely.

- Asking for a number format on a package with no style table is refused
  instead of being dropped without a word.

- A cell reference no column letter can spell no longer breaks the lookups
  around it. One `XFE1` in a sheet made every `cell()` call in it throw, while
  `cells()` returned the whole sheet.

- Writing to a sheet whose part is missing from the package is refused. The
  edit was accepted, reported by `cell()`, and saved nowhere.

- A lone surrogate is refused rather than written. Encoding replaced it with
  U+FFFD, so the cell held one string in memory and a different one in the
  file, with nothing to say so. U+FFFE and U+FFFF are refused too: both survive
  our own reader and neither is legal XML, so the file was one a stricter
  parser would reject.

- Overwriting the cell an array formula or a data table spills from is refused,
  as it already was for a shared formula. The rest of the range holds cached
  values and no formula, so replacing the anchor left them owned by nothing.

- Date format detection follows the format code sections. A code's fourth
  section applies only to text, so `#,##0;-#,##0;0;"due "mmm` is a number
  format, and the character after `*` or `_` is a literal like the one after
  `\` already was.
- Dropping the calculation chain drops the relationship naming it. A
  relationship pointing at a part that is no longer in the package is invalid,
  and it is the kind of thing Excel offers to repair.
- A written string is no longer pointed at a shared string built from formatted
  runs, or one carrying a phonetic guide. Both read as the same text, so the
  cell silently inherited the formatting of an entry it merely matched.
- A number too large to be a date, in a cell that carries a date format, reads
  as the number it is. It threw, and because the throw came out of the iterator
  the whole sheet became unreadable. Excel shows such a cell as `###`, so the
  file was legal and we were the ones refusing it.
- A `Date` after 9999 is refused by `set`, naming the cell. It was accepted,
  written, and unreadable afterwards, because only the lower bound was checked.
- An edit refused by the number format it asks for records nothing. The edit was
  queued before the format was resolved, so `set` threw, `cell()` reported the
  old value as though the write had been rejected, and `toBytes()` wrote the new
  one anyway.
- A requested number format no longer strips the font, fill and border of the
  cell it is applied to.
- `numFmts` keeps its `count` in step with its children. A mismatch makes Excel
  treat the file as unreadable and offer to repair it.
- Carriage returns are escaped. Written raw, XML parsing folds them into bare
  line feeds, so `a\r\nb` came back as `a\nb`.
- Shared strings, cell formats and the recalculation flag are written the way
  the file writes them, rather than assuming no namespace prefix and double
  quotes.
- A row that omits its `r` is numbered as the row after the one before it. The
  write path counted rows seen instead, so once any row declared a number, an
  edit to a bare row appended a second row with the same number.
- Style lookups use the canonical reference, so a file spelling a cell `a1` or
  `$A$1` no longer loses that cell's formatting on an edit.
