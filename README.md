# xlsx-corpus-harness

Measures whether an XLSX library can read a file and write it back without
damaging it. Fixed corpus, so a number that moves means the library changed.

## Usage

```bash
npm install
npm run corpus     # build the synthetic files
npm run harness    # measure every registered adapter
npm run harness -- quirks/shared-formula   # filter to matching paths
```

## What it measures

For each file, per adapter:

| Signal | Meaning |
| --- | --- |
| `parts dropped` | ZIP parts present on input, absent on output |
| `features degraded` | Markup constructs whose count fell: conditional formatting, data validation, merged cells, hyperlinks, formulas, defined names |
| `cell values lost/altered` | Values or per-cell formatting that changed across the round trip |

Volatile parts (`docProps/core.xml`, `docProps/app.xml`, `xl/calcChain.xml`) and
directory entries are excluded — they differ on every write and are not loss.

Style registry counts are not measured. A writer that prunes an unused number
format or dedupes two identical fonts has lost nothing. Formatting is compared
per cell instead, via the fingerprint on `CellValue.style`.

## Corpus

- `corpus/real/` — 60 files from Apache POI's test data, committed and pinned.
  Produced by real spreadsheet applications, which is what makes them useful.
  See `PROVENANCE.md`.
- `corpus/quirks/` — hand-built files covering legal-but-unusual constructs:
  inline strings, omitted cell references, rows out of order, a lying
  `dimension`, the 1904 epoch, the 1900 leap-year bug, shared formulas, columns
  past Z, XML entities, boolean and error cells.
- `corpus/generated/` — workbooks built with ExcelJS. Weak by construction: a
  library reading its own output proves little. They are a floor, not a bar.
- `corpus/manual/` — drop your own files here. Gitignored.

## Adding a library

Implement `Adapter` from `src/harness/types.ts` and register it in
`src/harness/run.ts`:

```ts
export interface Adapter {
  name: string
  roundTrip(bytes: Uint8Array): Promise<Uint8Array>
  values(bytes: Uint8Array): Promise<SheetValues[]>
}
```

## Baseline

`exceljs@4.4.0` over the 72 committed files: 25 clean, 45 lossy, 2 unreadable.

Damage from opening and saving with no edits:

| Feature | Files affected |
| --- | --- |
| charts | 12 |
| drawings | 12 |
| definedNames | 8 |
| colWidths | 4 |
| pivotTables | 3 |
| pivotCaches | 3 |
| customXml | 2 |
| mergedCells | 1 |
| hyperlinks | 1 |

Unreadable:

- `quirks/missing-cell-refs.xlsx` — `Invalid row number in model`. The `r`
  attribute on `<row>` and `<c>` is optional in ECMA-376.
- `real/dataValidationTableRange.xlsx` — `Cannot set properties of undefined
  (setting 'filterButton')`.

The synthetic files pass almost entirely while the real ones fail 45 of 60,
which is the argument for keeping real files in the corpus.
