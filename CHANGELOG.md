# Changelog

Follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and semantic
versioning. Nothing has been published yet, so everything below is unreleased
and the API is still free to change.

## Unreleased

### Added

- `set(reference, { formula })` writes a formula. Written without a cached
  result, so the cell reads back empty until a spreadsheet application opens the
  file and calculates it. The workbook is marked for recalculation.
- `set(reference, value, { numberFormat })` chooses a number format code,
  reusing a format the file already has where one matches.
- `npm run external` converts every real test file with LibreOffice, before and
  after an edit, to check the output against an implementation that is not this
  one.
- The round-trip harness measures an edited file as well as an untouched one,
  and reports parts rewritten separately from data lost.

### Changed

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

- Dropping the calculation chain drops the relationship naming it. A
  relationship pointing at a part that is no longer in the package is invalid,
  and it is the kind of thing Excel offers to repair.
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
