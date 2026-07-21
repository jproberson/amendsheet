# xlsxdoc

Read and edit `.xlsx` files in Node and the browser.

Anything the library doesn't interpret gets written back byte for byte. Charts,
pivot tables, drawings and macros all survive a read and a save, so editing one
cell won't quietly throw away the rest of the workbook.

One dependency, `fflate`, for the ZIP container. There is no transitive tree to
audit.

## Use

```ts
import { readWorkbook } from 'xlsxdoc'

const workbook = readWorkbook(bytes)

for (const cell of workbook.sheets[0].cells()) {
  console.log(cell.reference, cell.value)
}
```

Editing:

```ts
const workbook = readWorkbook(bytes)

workbook.sheets[0].set('B7', 42)
workbook.sheets[0].set('C1', 'a new cell')
workbook.sheets[0].set('D2', new Date('2024-01-01'))

await writeFile('out.xlsx', workbook.toBytes())
```

`set` takes a number, string, boolean, `Date`, or `null` to clear a cell. If the
cell or its row isn't there yet, both get created. The style of a replaced cell
is kept, so a `Date` written into a cell with no date format shows up as a
number.

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

## Not done yet

- Text is written as an inline string instead of going into the shared string
  table, so repeated text takes more bytes than it needs to.
- Writing isn't streamed. A sheet is patched as one string.
- Charts, pivot tables and drawings are preserved but never created.

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
