# AmendSheet

Amend `.xlsx` files. Everything you didn't touch stays exactly as it was.

Reads and edits spreadsheets in Node and the browser. Anything the library
doesn't interpret gets written back byte for byte, so charts, pivot tables,
drawings and macros all survive a read and a save. Changing one cell won't
quietly throw away the rest of the workbook.

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

sheet.set('B7', 42)
sheet.set('C1', 'a new cell')
sheet.set('D2', new Date('2024-01-01'))
sheet.set('E1', null) // clears the value, keeps the formatting
sheet.set('F9', { formula: 'SUM(F1:F8)' })
sheet.set('G1', 1234.5, { format: '"$"#,##0.00' })

const bytes = workbook.toBytes() // synchronous
await writeFile('out.xlsx', bytes)
```

`set` takes a number, string, boolean, `Date`, `{ formula }`, or `null` to clear
a cell. Pass `{ format }` to choose a number format code; without one the cell
keeps whatever formatting it had. If the
cell or its row isn't there yet, both get created, and the declared dimension
grows to cover them. The style of a replaced cell is kept, so its formatting
survives the edit. Writing a `Date` into a cell with no date format applies one,
reusing a format the file already has where possible.

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

Dates are calendar dates, not instants. `new Date(2024, 0, 1)` writes the serial
for that day, and reading it back gives a `Date` whose `getDate()` is 1, in
whatever timezone the code runs in. Build dates the ordinary way rather than
with `Date.UTC`.

Reading a formula gives you `cell.formula`, the expression without the leading
`=`. A cell that follows a shared formula reports an empty string, since the
expression lives on the master cell.

Write one with `{ formula }` rather than a string beginning with `=`, so text
that happens to start with `=` stays text. Nothing here computes a result, so
the cell is written without one and the workbook is marked for recalculation.
Until something opens it, that cell reads back with a value of `kind: 'empty'`
and its expression in `cell.formula`.

## Not done yet

- Overwriting the cell that defines a shared formula is refused rather than
  breaking the cells that follow it.
- Nothing evaluates formulas, so a written one has no value until a spreadsheet
  application opens the file.
- Writing isn't streamed. A sheet is patched as one string.
- Charts, pivot tables and drawings are preserved but never created.
- Nothing reads or writes cell formatting beyond number formats, so fonts,
  fills and borders can be preserved but not set.

## Speed

`node scripts/bench.mjs` writes into a sheet of the given size. Against 10,000
rows, on an M-series laptop:

```
write 10000 cells                             62 ms
write 10000 cells, reading between each      102 ms
write 10000 dates                            108 ms
append 10000 new rows                         88 ms
```

Each is linear in the number of edits. Reading between writes used to be
quadratic, and the same run took 7 minutes 46 seconds.

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

## Round-trip harness

`npm run harness` reads every test file, writes it straight back out, and reports
what changed: missing ZIP parts, markup features whose count dropped, and cell
values that differ. Nothing changes across the 72 committed files, and that's the
regression gate for the preservation guarantee.

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
