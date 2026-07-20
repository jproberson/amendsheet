# xlsxdoc

Reads `.xlsx` files in Node and the browser. Parts it does not interpret —
charts, pivot tables, drawings, macros — are written back byte for byte, so
opening a file and saving it does not damage it.

Zero dependencies except [fflate](https://github.com/101arrowz/fflate) for the
ZIP container.

Status: read path only. Editing and writing cell values are not implemented yet.

## Use

```ts
import { readWorkbook } from 'xlsxdoc'

const workbook = readWorkbook(bytes)

for (const sheet of workbook.sheets) {
  for (const cell of sheet.cells()) {
    console.log(cell.reference, cell.value)
  }
}

const saved = workbook.toBytes()
```

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
coverage thresholds, and prints the harness numbers. Run it before every commit.

## Measurements

Round-trip fidelity over the 72 committed corpus files — read a file, write it
straight back, compare every ZIP part, markup feature, and cell value:

```
xlsxdoc            72 clean  |   0 lossy  |  0 failed to process
exceljs@4.4.0      25 clean  |  45 lossy  |  2 failed to process
```

Opening and saving with ExcelJS damages charts in 12 files, drawings in 12,
defined names in 8, column widths in 4, and pivot tables in 3. Two files it
cannot read at all: one omits the optional `r` attribute on `<row>` and `<c>`,
which is legal, and one trips an internal error on a data validation range.

This comparison is not symmetric. ExcelJS rebuilds a file from its object model;
xlsxdoc currently writes the parts it read. The interesting measurement is still
ahead, when writing modified cells lands.

Reproduce with `npm run harness`. Add a library by implementing `Adapter` in
`src/harness/types.ts` and registering it in `src/harness/run.ts`.

## Corpus

- `corpus/real/` — 60 files from Apache POI's test data, committed and pinned to
  a commit. Written by real spreadsheet applications. See `PROVENANCE.md`.
- `corpus/quirks/` — hand-built files covering legal but unusual constructs:
  inline strings, omitted cell references, rows out of order, a lying
  `dimension`, the 1904 epoch, the 1900 leap-year bug, shared formulas, columns
  past Z, XML entities, boolean and error cells.
- `corpus/generated/` — workbooks built with ExcelJS. Weak by construction: a
  library reading its own output proves little.
- `corpus/manual/` — drop your own files here. Gitignored.

`npm run corpus` builds the generated and quirk files. `npm run corpus:real`
pulls more from the pinned POI commit.
