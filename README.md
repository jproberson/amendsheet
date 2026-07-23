# AmendSheet

Amend `.xlsx` files. Everything you didn't touch stays exactly as it was.

Reads and edits spreadsheets in Node and the browser, checked in both. The
contents of any part the library doesn't interpret get written back byte for
byte, so charts, pivot tables, drawings and macros all survive a read and a
save. Changing one cell won't quietly throw away the rest of the workbook.

The ZIP container itself is rebuilt, so an untouched file comes back the same
document but not the same bytes: compression, timestamps and entry order are
the writer's own.

One dependency, `fflate`, for the ZIP container. There is no transitive tree to
audit.

## On AI

Heads up: this was built with heavy use of AI assistance, partly as a test of
what that produces when held to normal standards. Tests are written first,
`./verify.sh` gates every commit, and the test files are real spreadsheets this
library had no hand in creating.

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

const out = workbook.toBytes() // synchronous
await writeFile('out.xlsx', out)
```

`set` takes a number, string, boolean, `Date`, `{ formula }`, or `null` to clear
a cell. Pass `{ numberFormat }` to choose a format code; without one the cell
keeps whatever formatting it had. If the
cell or its row isn't there yet, both get created, and the declared dimension
grows to cover them. The style of a replaced cell is kept, so its formatting
survives the edit. Writing a `Date` into a cell with no date format applies one,
reusing a format the file already has where possible.

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
`color` and `name`, a font carries `strike`, `vertAlign` (`superscript` or
`subscript`) and `underline`, which is `true` for a single underline or one of
`double`, `singleAccounting`, `doubleAccounting`. A `fill` is a discriminated
union: `{ type: 'solid', color }` for a plain background, or `{ type: 'pattern',
pattern, color, background }` for one of the ECMA-376 pattern types such as
`lightGrid`, where `color` is the pattern's foreground. A
`border` sets sides by name, or `all` at once, merging onto the sides the cell
already has. An `alignment` places the text — `horizontal`, `vertical`,
`wrapText`, `textRotation` (0–180, or 255 to stack top to bottom) and `indent` —
and merges the same way, so setting `wrapText` leaves a horizontal choice alone.
A `protection` sets `locked` and `hidden`, which take effect once the worksheet
itself is protected. Colours are `RRGGBB` or `AARRGGBB` hex.

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
  library cannot use; every other code is about the file, or the runtime reading
  it, not a value you passed.
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
- Charts, pivot tables and drawings are preserved but never created.
- `cell` exposes a cell's `font`, `fill`, `border` and `alignment`, but only the
  parts this library models. Anything else on the cell format — protection, a
  gradient or pattern fill, a themed colour — is preserved in the file but not
  reported.
- A table grows to include a cell written directly below or to the right of it,
  the way Excel would, adding a column when it grows sideways. Other ranges that
  name cells are still copied, not adjusted: chart ranges, defined names and
  conditional formatting keep the extent they had.
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
