# xlsxdoc

Read and edit `.xlsx` files in Node and the browser.

Parts it does not interpret — charts, pivot tables, drawings, macros — are
written back byte for byte, so a file survives being opened, changed and saved
with everything else intact.

One dependency: [fflate](https://github.com/101arrowz/fflate), for the ZIP
container. Nothing else, so there is no transitive tree to audit.

## Use

```ts
import { readWorkbook } from 'xlsxdoc'

const workbook = readWorkbook(bytes)

for (const cell of workbook.sheets[0].cells()) {
  console.log(cell.reference, cell.value)
}
```

Editing a file leaves everything you did not touch exactly as it was:

```ts
const workbook = readWorkbook(bytes)

workbook.sheets[0].set('B7', 42)
workbook.sheets[0].set('C1', 'a new cell')
workbook.sheets[0].set('D2', new Date('2024-01-01'))

await writeFile('out.xlsx', workbook.toBytes())
```

`set` accepts a number, string, boolean, `Date`, or `null` to clear a cell. It
creates the cell and its row if they are not there yet, and carries over the
style of a cell it replaces — which means a `Date` written into a cell with no
date format will display as a number.

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

A date is a number wearing a date number format, so `kind: 'date'` is produced
by resolving the cell's style. The original `serial` is kept alongside the
`Date` so the stored value can be written back unchanged.

## Not done yet

- Text is written as an inline string rather than added to the shared string
  table, so repeated text costs more bytes than it should.
- Writing is not streamed: a sheet is patched as one string.
- Nothing writes charts, pivot tables or drawings. They are preserved, not
  created.

## Layout

```
src/lib/       the library
src/harness/   round-trip measurement rig
src/adapters/  library implementations the harness measures
src/corpus/    corpus generation and fetching
corpus/        the files the harness runs against
```

## Verify

```bash
./verify.sh
```

Formats, lints, typechecks, greps for banned constructs, runs the tests with
coverage thresholds, and checks the built package. Run it before every commit.

## Round-trip harness

`npm run harness` reads every file in the corpus, writes it straight back, and
reports anything that changed: ZIP parts that went missing, markup features whose
count fell, and cell values that differ. It is the regression gate for the
preservation guarantee — over the 72 committed files, nothing changes.

## Corpus

- `corpus/real/` — 60 files from Apache POI's test data, committed and pinned to
  a commit. Written by real spreadsheet applications. See `PROVENANCE.md`.
- `corpus/quirks/` — hand-built files covering legal but unusual constructs:
  inline strings, omitted cell references, rows out of order, a lying
  `dimension`, the 1904 epoch, the 1900 leap-year bug, shared formulas, columns
  past Z, XML entities, boolean and error cells.
- `corpus/generated/` — feature-rich workbooks built by another library, as a
  smoke test. Weaker than the real files, which no JavaScript library produced.
- `corpus/manual/` — drop your own files here. Gitignored.

`npm run corpus` builds the generated and quirk files. `npm run corpus:real`
pulls more from the pinned POI commit.
