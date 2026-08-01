# AmendSheet

Amend `.xlsx` files. Everything you didn't touch stays exactly as it was.

Reads, edits and creates spreadsheets in Node and the browser, checked in both. The
contents of any part the library doesn't interpret get written back byte for
byte, so charts, pivot tables, drawings and macros all survive a read and a
save. Changing one cell won't quietly throw away the rest of the workbook.

The ZIP container itself is rebuilt, so an untouched file comes back the same
document but not the same bytes: compression, timestamps and entry order are
the writer's own.

One dependency, `fflate`, for the ZIP container. There is no transitive tree to
audit.

## Fidelity

The claim is measured, not asserted. `npm run harness` reads 73 real files —
Apache POI's test data plus hand-built edge cases — writes each back, and
compares the result part by part and value by value. Nothing is lost and no
untouched part is rewritten, before an edit or after one.

For a reference point the same files go through `exceljs@4.4.0`, a mature and
much broader library: it keeps 25 of them intact and drops a part or a value in
46, most often a chart, a drawing or a defined name. That is the cost of
rebuilding the whole document on write, and it is the one thing this library sets
out not to do. The full table and the method are in
[`COMPARISON.md`](COMPARISON.md), regenerated with `npm run harness:doc`.

## On AI

Heads up: this was built with heavy use of AI assistance, partly as a test of
what that produces when held to normal standards. Tests are written first,
`./verify.sh` gates every commit, and the test files are real spreadsheets this
library had no hand in creating.

For an AI writing code against the library, [`llms.txt`](llms.txt) is the whole
API in one flat, example-dense file — load it and write correct code without
guessing. Its examples are typechecked against the exports on every `verify.sh`,
so they cannot drift.

## Use

```ts
import { readWorkbook } from 'amendsheet'

const workbook = readWorkbook(bytes)

const sheet = workbook.sheet('Summary') ?? workbook.sheets[0]

for (const cell of sheet.cells()) {
  console.log(cell.reference, cell.value)
}

sheet.cell('B7')?.value
```

`cells()` yields every cell the sheet stores, which includes cells holding only
formatting and cells that were cleared. Those arrive as `kind: 'empty'`, so
filter on `kind` if you only want content.

Editing:

```ts
const workbook = readWorkbook(bytes)
const sheet = workbook.sheet('Summary') ?? workbook.sheets[0]

sheet.set('B7', 42)
sheet.set('C1', 'a new cell')
sheet.set('D2', new Date('2024-01-01'))
sheet.set('E1', null) // clears the value, keeps the formatting
sheet.set('F9', { formula: 'SUM(F1:F8)' })
sheet.set('G1', 1234.5, { numberFormat: '"$"#,##0.00' })

// Write a block of values from a top-left cell, and read a range back.
sheet.setValues('A10', [
  ['name', 'qty'],
  ['apple', 3],
])
const block = sheet.getValues('A10:B11') // CellValue[][]

const out = workbook.toBytes() // synchronous
await writeFile('out.xlsx', out)
```

`set` takes a number, string, boolean, `Date`, `{ formula }`, `{ runs }` for rich
text, or `null` to clear a cell. Pass `{ numberFormat }` to choose a format code;
without one the cell keeps whatever formatting it had. If the
cell or its row isn't there yet, both get created, and the declared dimension
grows to cover them. The style of a replaced cell is kept, so its formatting
survives the edit. Writing a `Date` into a cell with no date format applies one,
reusing a format the file already has where possible.

Rich text lets one cell mix fonts. Pass `{ runs }`, each run a stretch of text
with an optional `font` that inherits the cell's where absent; it is written as an
inline string. Reading gives the whole string back as `cell.value` and the runs as
`cell.richText`, so code that only wants the text can ignore the formatting.

```ts
const workbook = readWorkbook(bytes)
const sheet = workbook.sheet('Summary') ?? workbook.sheets[0]

sheet.set('A1', {
  runs: [{ text: 'Total: ' }, { text: '42', font: { bold: true, color: 'FF0000' } }],
})
const cell = sheet.cell('A1')
cell?.value // { kind: 'text', value: 'Total: 42' }
cell?.richText?.runs // the two runs, the second bold and red
```

Formatting goes through the same options, and composes into one cell format:

```ts
const workbook = readWorkbook(bytes)
const sheet = workbook.sheet('Summary') ?? workbook.sheets[0]

sheet.set('A1', 1234.5, {
  numberFormat: '"$"#,##0.00',
  font: { bold: true, color: 'FF0000' },
  fill: { type: 'solid', color: 'FFFF00' },
  border: { all: { style: 'thin' } },
  alignment: { horizontal: 'center', wrapText: true },
  protection: { locked: false },
})

// format() changes only the formatting, so a formula cell keeps its expression.
sheet.format('B2', { font: { italic: true }, border: { bottom: { style: 'medium' } } })
```

A `font` merges onto the cell's current one, so `{ font: { bold: true } }` adds
bold without disturbing its size or colour. Besides `bold`, `italic`, `size`,
`color` and `name`, a font carries `strike`, `verticalAlign` (`superscript` or
`subscript`) and `underline`, which is `true` for a single underline or one of
`double`, `singleAccounting`, `doubleAccounting`. A `fill` is a discriminated
union: `{ type: 'solid', color }` for a plain background, or `{ type: 'pattern',
pattern, color, background }` for one of the ECMA-376 pattern types such as
`lightGrid`, where `color` is the pattern's foreground. A
`border` sets sides by name, or `all` at once, merging onto the sides the cell
already has; a `diagonal` side — `{ style, color?, up?, down? }` — draws a line
across the cell, on the up diagonal, the down one, or both, defaulting to down.
An `alignment` places the text — `horizontal`, `vertical`,
`wrapText`, `textRotation` (0–180, or 255 to stack top to bottom) and `indent` —
and merges the same way, so setting `wrapText` leaves a horizontal choice alone.
A `protection` sets `locked` and `hidden`, which take effect once the worksheet
itself is protected. A colour is an `RRGGBB` or `AARRGGBB` hex string, a theme
reference `{ theme, tint }` (an index into the workbook's colour scheme; `tint`
lightens or darkens it, -1 to 1), or an indexed palette entry `{ indexed }`. A
theme or indexed colour is read back and written as the reference it is, not
resolved to the hex it displays as, so editing a cell keeps the theme colour its
font already carried rather than dropping it.

When you do want the displayed hex, `workbook.resolveColor(color)` returns the
8-digit `AARRGGBB` a reference resolves to, reading the workbook's theme and
applying any tint:

```ts
const workbook = readWorkbook(bytes)
const cell = workbook.sheets[0]?.cell('B2')
const shown = cell?.font?.color ? workbook.resolveColor(cell.font.color) : undefined
```

It returns `undefined` for a colour with no fixed value — a system indexed
colour, or a theme slot the workbook's theme does not define. A plain hex passes
through; a tinted theme colour is resolved in the colour space Excel uses and
matches its shown value to within one unit per channel.

`worksheet.protect()` turns that worksheet protection on — it is what makes a
cell's `locked` and `hidden` bite. With no argument it matches Excel's Protect
Sheet default: every cell locked, formatting, inserting, deleting, sorting and
filtering barred, and selecting cells still allowed. Pass options to name the
actions that stay permitted, each `true` to allow it:

```ts
const sheet = readWorkbook(bytes).sheets[0]
sheet?.protect({ formatCells: true, sort: true, selectLockedCells: false })
```

It replaces any protection the sheet already had. Passwords are not written, so
protection guards against accidental edits, not a determined one.
`worksheet.protection` reads the protection back in the same shape, or is
undefined when the sheet is not protected.

`worksheet.merge('A1:B2')` merges a range, joining any merges the sheet already
has. Excel shows only the top-left cell; the others keep their values, since a
merge does not clear them, and a write to one is refused the way it already was
for a merge the file came with.

`worksheet.addImage(bytes, 'B2:E10')` embeds a picture spanning a cell (or range),
its type read from the image's own bytes (PNG, JPEG or GIF). It joins the sheet's
drawing, one being created if there is none, so several images share the one part.

`worksheet.setRowHeight(1, 30)` sizes a row in points, and
`worksheet.setColumnWidth('A', 24)` sizes a column in Excel's width units,
splitting a `cols` range that spans more than the one column so the rest keeps
its own width. `worksheet.hideRow(3)` and `worksheet.hideColumn('C')` hide a
line, keeping any height or width it also has.

`worksheet.insertRows(before, count?)` inserts blank rows before a row number,
and `insertColumns(before, count?)` inserts blank columns before a column letter;
`deleteRows(from, count?)` and `deleteColumns(from, count?)` take them out.

```ts
const sheet = readWorkbook(bytes).sheets[0]

sheet.insertRows(3, 2)    // two blank rows above row 3
sheet.deleteColumns('B')  // remove column B, pulling C onward left
```

Every reference that points into the moved lines moves with them — a formula
anywhere in the workbook, a shared formula's range, merges, the dimension,
filters, conditional formats, validations, hyperlinks and defined names. A cell
set earlier in the same session lands first, so it rides along. A deletion turns
a reference to a removed cell into `#REF!` and shrinks a range that only
overlapped it. An insert that would push a line off the sheet is refused, as is a
deletion that would collapse a whole merge, filter, format or shared formula, and
either on a sheet carrying a table, whose stored range is not adjusted yet.

`worksheet.freeze('B2')` freezes the rows above and columns left of a cell, and
`worksheet.autoFilter('A1:D100')` sets the filter over a range, replacing any the
sheet already has. `worksheet.tabColor('FF0000')` colours the sheet tab, taking a
6- or 8-digit hex and replacing any colour the tab already had.
`worksheet.showGridlines(false)` and `worksheet.showHeadings(false)` toggle the
gridlines and the row/column headings, and `worksheet.zoom(85)` sets the zoom as a
whole percentage. All three land on the sheet's view, alongside a freeze rather
than replacing it. `worksheet.groupRows(2, 5)` and `worksheet.groupColumns('B', 'D')`
set an outline level over a range so a reader shows a collapsible band; a third
argument nests a deeper level (1 to 7) inside a shallower one.

`worksheet.validate(range, rule)` adds a data validation. The rule is a `list` — a
dropdown whose values may not hold a comma, since an inline list separates them
with one — or a `whole`/`decimal` numeric comparison:

```ts
const workbook = readWorkbook(bytes)
workbook.sheets[0]?.validate('B2:B10', { list: ['Yes', 'No', 'Maybe'] })
workbook.sheets[0]?.validate('C2:C10', { whole: { between: [1, 100] } })
workbook.sheets[0]?.validate('D2:D10', { decimal: { greaterThan: 0 } })
```

A comparison is one of `between`, `notBetween`, `equal`, `notEqual`,
`greaterThan`, `lessThan`, `greaterThanOrEqual` or `lessThanOrEqual`, each bound a
finite number. A validation joins any the sheet already carries and lands in the
schema's place, so a reader takes the file without offering to repair it.

`worksheet.conditionalFormat(range, rule)` adds a conditional format. A
`colorScale` grades cells between two colours, or three with a `mid`; a `cellIs`
fills the cells whose value meets a comparison:

```ts
const workbook = readWorkbook(bytes)
workbook.sheets[0]?.conditionalFormat('A1:A20', { colorScale: { min: 'FFFFFF', max: 'FF0000' } })
workbook.sheets[0]?.conditionalFormat('B1:B20', {
  colorScale: { min: 'F8696B', mid: 'FFEB84', max: '63BE7B' },
})
workbook.sheets[0]?.conditionalFormat('C1:C20', {
  cellIs: { when: { greaterThan: 100 }, fill: 'FFFF00' },
})
```

The `when` of a `cellIs` is any of the comparisons `validate` takes; its `fill` is
written as a differential format in `styles.xml`. A `{ dataBar: { color } }` draws
a bar in each cell, scaled between the range's smallest and largest values. A rule
outranks any the sheet already has, so a new format wins where they overlap.

`worksheet.comment(reference, text)` attaches a comment to a cell, and
`cell.comment` reads one back. A comment's text is written into a comments part
wired to the sheet; the box's visual shape (the legacy VML drawing) is not
written, so a reader that needs it may place the note itself.

```ts
const workbook = readWorkbook(bytes)
workbook.sheets[0]?.comment('B2', 'Double-check this figure')
const note = workbook.sheets[0]?.cell('B2')?.comment
```

Adding a comment to a sheet that already has some is refused with
`unsupported-edit`, since merging into its part would rebuild the rich text the
existing comments hold as plain words.

`workbook.defineName('TaxRate', 'Sheet1!$B$1')` defines a global named range, and
`workbook.definedNames` reads them back. The same three — `defineName`,
`removeDefinedName` and `definedNames` — sit on a `worksheet` for names scoped to
one sheet, where two sheets may each hold a name of the same spelling. A sheet's
print area is one such scoped name: `worksheet.setPrintArea('A1:J26')` sets it,
`worksheet.printArea` reads it back as a plain range, and `clearPrintArea()`
removes it. `worksheet.link('A1', { url: 'https://example.com' })` links a cell out to a
URL, and `link('A1', { location: 'Sheet2!A1' })` links within the workbook to a
cell or a defined name — both take an optional `tooltip`, replace any link the
cell had, and move with an inserted or deleted line.

A workbook does not have to come from a file. `createWorkbook()` starts a blank
one with a single empty sheet named `Sheet1`, or the name you pass, ready for the
same edit API. `workbook.addSheet('Data')` adds an empty sheet and returns it,
`worksheet.rename('Report')` renames one, and `worksheet.remove()` drops it,
refusing to remove the last. `workbook.copySheet('Template', 'March')` duplicates a
sheet and returns the copy, bringing its cells, formatting, formulas, merges and
comments across into parts of its own; a sheet carrying a table, drawing or pivot
table is refused, since those need names, ids or media reworked to stay valid.
Creating goes through the same path a read workbook does, so a created sheet takes
every edit above.

```ts
import { createWorkbook } from 'amendsheet'

const workbook = createWorkbook('Budget')
workbook.sheets[0].set('A1', 'Item')
workbook.addSheet('Notes').set('A1', 'draft')
const out = workbook.toBytes()
```

`createWorkbookFromCsv(text, options?)` builds one from delimited text — fields are
text unless `parseNumbers` is set — and `worksheet.toCsv(options?)` prints a sheet
back out, RFC 4180 style, from `A1` to the furthest cell that holds anything.

An edit the format cannot hold is refused by `set` itself, with an `XlsxError`
naming the cell. `NaN`, an infinity, a character XML has no way to encode, a
date outside the workbook's epoch, and overwriting the cell that defines a
shared formula are all refused this way. A refused edit is not recorded, so the
rest of the batch still writes and `cell()` never reports a value the file is
not going to receive.

`cell.value` is a discriminated union:

```ts
type CellValue =
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'empty' }
  | { kind: 'date'; value: Date; serial: number }
```

There is no date type in the file format. A date is a number with a date number
format applied, so `kind: 'date'` comes from resolving the cell's style. The
`serial` is kept so the stored value can go back unchanged.

Which format codes count as dates is a heuristic over the code, and it is not
frozen: a release may start or stop calling a given format a date, and a cell
can move between `number` and `date` without the type changing to warn you. Both
carry the same stored double, in `serial` or `value`, so read it by kind:

```ts
const stored =
  value.kind === 'date' ? value.serial : value.kind === 'number' ? value.value : undefined
```

If you need a stable classification, read `cell.numberFormat` and decide for
yourself.

Dates are calendar dates, not instants. `new Date(2024, 0, 1)` writes the serial
for that day, and reading it back gives a `Date` whose `getDate()` is 1, in
whatever timezone the code runs in. Build dates the ordinary way rather than
with `Date.UTC`.

Reading a formula gives you `cell.formula`, which is either
`{ kind: 'expression', expression }` — the source without the leading `=` — or
`{ kind: 'shared', master }` for a cell that follows a shared formula, naming
the cell the expression is stored on. A shared formula is written once and
reused down a column, so the followers hold a cached value and nothing else.

Write one with `{ formula }` rather than a string beginning with `=`, so text
that happens to start with `=` stays text. Nothing here computes a result, so
the cell is written without one and the workbook is marked for recalculation.
Until something opens it, that cell reads back with a value of `kind: 'empty'`
and its expression in `cell.formula`.

## Compatibility

What a minor release is allowed to do, so you know which branches are safe:

- **`XlsxErrorCode` is open.** New codes arrive in minor releases, because there
  is no knowing today every way an xlsx can be malformed. Switch with a default.
  An existing code will not change meaning without a major version.
  `bad-reference` and `unwritable-value` mean the caller passed something the
  library cannot use, and `unsupported-edit` means an edit is well formed but not
  one the library performs yet; every other code is about the file, or the
  runtime reading it, not anything you passed.
- **`kind: 'date'` is not a frozen classification.** See above.
- **A `Date` written and read back is the same date.** A serial another
  application wrote can carry finer resolution than a millisecond, which is all
  a `Date` holds.

## Not done yet

- Overwriting the anchor of a shared formula, an array formula or a data table
  is refused rather than breaking the cells that spill from it. Writing into a
  merged cell that is not the anchor is refused for the same reason: a value
  there would never show.
- A digital signature survives as a part but any edit invalidates it, since the
  package is rebuilt and the bytes it signed change.
- Nothing evaluates formulas, so a written one has no value until a spreadsheet
  application opens the file.
- A part the edit does not touch is copied through still compressed, so a large
  workbook is not fully decompressed to change one cell. A sheet is read and
  patched in its bytes rather than decoded to one string, so the ceiling on a
  single sheet is the largest buffer this runtime allocates rather than V8's
  ~512MB string limit; past that a part is refused with `part-too-large`.
  Nothing streams a sheet through in chunks, though, so it is still held whole
  in memory.
- Charts and pivot tables are preserved but never created. A drawing is preserved,
  and `worksheet.addImage` adds a picture to one; other drawing objects are not
  authored.
- `cell` exposes a cell's `font`, `fill`, `border` and `alignment`, but only the
  parts this library models. A fill is solid, a pattern or a gradient, and each is
  one `set` can write as well as read; `set` writes a gradient as a linear one, so
  a path gradient read from a file comes back linear if it is written again. A
  theme or indexed colour is reported as the reference it is; `workbook.resolveColor`
  turns one into the hex it displays as.
- A table grows to include a cell written directly below or to the right of it,
  the way Excel would, adding a column when it grows sideways. Inserting or
  deleting rows and columns moves the references that name cells — formulas,
  merges, filters, conditional formats, hyperlinks and defined names. It also
  moves a table's own stored range and a comment's cell — the note's box in the
  legacy drawing moves with it, and a note whose cell a deletion removes is
  dropped. A table below or beside the edit shifts with it; a column inserted
  inside it gains a fresh column and header, and a deletion that clips it shrinks
  it, dropping the column entries it cut out. Only a deletion that would take every
  one of a table's columns, or a row deletion that would take its header row, is
  refused with `unsupported-edit`. A drawing moves too — a picture, shape, diagram
  or OLE object follows the edit by its cell anchor (one a deletion covers entirely
  is dropped), and a chart, the one drawing object that references cells, also has
  its series, category and title references shifted wherever the chart part lives,
  so a chart plotting the edited sheet from another one is caught. A pivot table
  moves too — its location follows the edit, and its cache's source range shifts
  when the sheet it reads is the one changed, wherever the pivot sits. Only a
  drawing whose part cannot be read, so its anchors cannot be moved, is refused.
- Chartsheets and dialogsheets aren't listed in `sheets`, since they hold no
  cells. They're still written back untouched.

## Speed

`node scripts/bench.mjs` writes into a sheet of the given size, in the four
shapes that have caused trouble: replacing cells, appending rows, writing dates,
and reading between every write.

Each is linear in the number of edits, and tens of milliseconds against 10,000
rows on a laptop. Reading between writes used to be quadratic; the same run took
7 minutes 46 seconds.

## Layout

```
src/lib/        the library
src/harness/    round-trip measurement
src/adapters/   libraries the harness measures
src/fixtures/   builds and fetches test files
fixtures/       the test files themselves
```

## Verify

```bash
./verify.sh
```

Formats, lints, typechecks, greps for banned constructs, runs the tests with
coverage thresholds, and checks the built package. Run it before every commit.

## How the tests are built

Coverage says a line ran, not that its output was right, so three things sit on
top of the ordinary tests.

**Invariants over fragments.** `src/testing/invariants.ts` holds the assertions
every write must satisfy: the sheet is well formed, cells sit inside rows, no
reference appears twice, and no part outside the edited sheet changed. Checking
for a substring passes on output that contains the right fragment inside a
broken document.

**Properties over examples.** `properties.test.ts` generates edits against all
60 real files and asserts what must hold for any of them: the sheet still
parses, edited cells read back as written, untouched cells are untouched, and
the order edits were made in does not matter. The seed is fixed so a failure
reproduces, and a failing case is shrunk to the fewest edits that still fail.
Hand-written tests only cover failure modes somebody already imagined.

**Mutation testing.** `npm run mutate` breaks the library one edit at a time and
checks that some test notices. It is slow, so it is not part of `verify.sh`.
A survivor is either a real gap or a mutation that changes nothing; both need
reading.

## Checking against something that isn't us

Every check above validates against this library's own reader, so a file we
write wrongly and read back wrongly looks right everywhere. `npm run external`
writes a value into each of the 60 real files and converts the result with
LibreOffice, which is a different OOXML implementation with its own opinions
about what is valid.

Each file is converted before and after the edit. A fixture LibreOffice can't
open to begin with proves nothing, so only one that opened before and fails
after counts. Currently 60 of 60 open after editing and LibreOffice finds the
written value in all of them.

It needs LibreOffice installed and takes a few minutes, so it isn't part of
`./verify.sh`.

`npm run browser` is the other outside check. It bundles the built library with
fflate, then reads a fixture, edits a cell and writes it back inside a headless
Chrome, confirming the code runs where no Node API exists. It drives Chrome over
the DevTools protocol directly, so it adds no dependency, and `./verify.sh` runs
it, skipping cleanly on a machine with no Chrome.

## Round-trip harness

`npm run harness` reads every test file and reports what changed: missing ZIP
parts, markup features whose count dropped, and cell values that differ. It runs
twice. Once writing the file straight back out, which measures the container and
nothing else. Then again after writing a cell past the last row in use, which is
the measurement that matches the claim — every existing cell has to come back
unchanged, and the only parts allowed to differ are the sheet the edit landed
in, the four parts an edit legitimately rewrites (`styles.xml`,
`sharedStrings.xml`, `workbook.xml`, `[Content_Types].xml`), and
`calcChain.xml`, which is deleted so the reader recomputes the values a formula
edit invalidates.

Nothing is lost or rewritten across the 73 committed files in either pass, and
both gate `./verify.sh`.

## Fixtures

- `fixtures/real/` has 60 files from Apache POI's test data, pinned to a commit.
  Real spreadsheet applications wrote these. See `PROVENANCE.md`.
- `fixtures/quirks/` holds hand-built files for legal but unusual constructs:
  inline strings, missing cell references, rows out of order, a `dimension` that
  lies, the 1904 epoch, the 1900 leap-year bug, shared formulas, columns past Z,
  XML entities, boolean and error cells.
- `fixtures/generated/` has feature-rich workbooks from another library, used as
  a smoke test. Weaker than the real files, since no JavaScript library wrote
  those.
- `fixtures/manual/` is for your own files. Gitignored.

`npm run fixtures` builds the generated and quirk files. `npm run fixtures:real`
pulls more from the pinned POI commit.

## License

MIT, see `LICENSE`.

The files in `fixtures/real/` come from Apache POI and stay under Apache-2.0,
with that project's `LICENSE` and `NOTICE` kept alongside them. They are test
data only. The published package contains just `dist/`, so none of it ships.
