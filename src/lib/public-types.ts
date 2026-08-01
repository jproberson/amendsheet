import type { CellAddress } from './reference.js'
import type {
  Alignment,
  BorderFormat,
  CellProtection,
  Color,
  FillFormat,
  FontFormat,
  ReadFill,
} from './styles-writer.js'
import type { CellInput, RichText, RichTextRun } from './cell-input.js'
import type { SheetProtection } from './patch.js'
import type { SheetState } from './workbook.js'
import type { Hyperlink } from './hyperlinks.js'
import type { DocumentProperties } from './document-properties.js'
import type {
  HeaderFooter,
  HeaderFooterSection,
  PageBreaks,
  PageMargins,
  PageSetup,
  PrintOptions,
} from './page.js'
import type { PrintTitles } from './defined-names.js'

export type { Hyperlink }
export type { RichText, RichTextRun }
export type { DocumentProperties }
export type { PageSetup, PageMargins, HeaderFooter, HeaderFooterSection, PageBreaks, PrintOptions }
export type { PrintTitles }

export type CellValue =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: string }
  | { readonly kind: 'empty' }
  /**
   * The stored number is kept so a date can be written back unchanged. Which
   * formats count as dates is a heuristic and may change in any release, so a
   * cell can move between `number` and `date`; `serial` and `value` are the
   * same double, and reading whichever is there survives that.
   */
  | { readonly kind: 'date'; readonly value: Date; readonly serial: number }

/**
 * A shared formula stores its source once, on the cell that owns the range.
 * The dependents carry a cached value and no source of their own, which is why
 * this is a union rather than a string that is sometimes empty.
 */
export type CellFormula =
  /** Without the leading `=`. */
  | { readonly kind: 'expression'; readonly expression: string }
  /** `master` is absent when the sheet holds no cell owning that group. */
  | { readonly kind: 'shared'; readonly master?: string }

export interface Cell {
  readonly address: CellAddress
  /**
   * Canonical, so it always equals `formatReference(cell.address)`. The one
   * exception is an address no column letter can name, which a lenient read
   * accepts and which keeps the spelling the file gave it.
   */
  readonly reference: string
  readonly value: CellValue
  /**
   * The cell's string as formatted runs, present only when it carries more than
   * one run or a run with formatting of its own. `value` still holds the whole
   * string flattened, so a caller that only wants the text can ignore this.
   */
  readonly richText?: RichText
  readonly formula?: CellFormula
  readonly numberFormat?: string
  /** Absent when the cell uses the default font, so carries none of its own. */
  readonly font?: FontFormat
  /** Absent unless the cell has a fill. Every shape reported here `set` can write. */
  readonly fill?: ReadFill
  /** Absent unless at least one side has a border. */
  readonly border?: BorderFormat
  /** Absent unless the cell sets an alignment of its own. */
  readonly alignment?: Alignment
  /** Absent unless the cell sets a `locked` or `hidden` protection of its own. */
  readonly protection?: CellProtection
  /** The cell's comment text, when it has one. Joined from the note's runs. */
  readonly comment?: string
  /** Where the cell links, when it links: a URL out of the package or a
   * `location` within the workbook. As read from the file, like `comment`; a
   * pending `link` shows after `toBytes`. A range link is reported on its
   * top-left. */
  readonly hyperlink?: Hyperlink
}

export interface Worksheet {
  readonly name: string
  /** Renames the sheet. The name follows the same rules as `addSheet`. */
  rename(name: string): void
  /** Removes the sheet from the workbook. A workbook must keep at least one. */
  remove(): void
  /** Visible, hidden, or very hidden (hidden and not offered in Excel's unhide
   * list). Reflects a pending `setState`. */
  readonly state: SheetState
  /**
   * Sets the sheet's visibility. Refused when it would hide the workbook's only
   * visible sheet, which Excel will not open.
   */
  setState(state: SheetState): void
  /**
   * As the workbook part spells it, so a defined name or a part this library
   * does not interpret can be matched against the sheet it refers to.
   */
  readonly sheetId: string
  /**
   * The names scoped to this sheet, each mapped to what it refers to. Two sheets
   * can each hold a scoped name of the same spelling, which is why these live on
   * the sheet rather than the workbook. Reflects a pending `defineName`. The
   * built-in `_xlnm.*` names are left out; the print area has its own accessor.
   */
  readonly definedNames: ReadonlyMap<string, string>
  /**
   * Defines a name scoped to this sheet, replacing one of the same spelling on it.
   * The name follows Excel's rules, like the workbook-level `defineName`; a
   * built-in `_xlnm.*` name is refused, since those carry their own accessors.
   */
  defineName(name: string, refersTo: string): void
  /** Removes a name scoped to this sheet. A name it does not have is ignored. */
  removeDefinedName(name: string): void
  /**
   * The sheet's print area — the range or comma-joined ranges that print — with
   * the sheet qualifier and the `$` of the stored form dropped, like `A1:J26`, or
   * undefined when the sheet has none. Reflects a pending `setPrintArea`.
   */
  readonly printArea: string | undefined
  /**
   * Sets the sheet's print area to a range like `A1:J26`, replacing any it has. It
   * is stored as the built-in `_xlnm.Print_Area` name scoped to the sheet. Refuses
   * anything that is not an `A1:B2` range.
   */
  setPrintArea(range: string): void
  /** Removes the sheet's print area, if it has one. */
  clearPrintArea(): void
  /**
   * The rows and columns the sheet repeats on every printed page — `rows` like
   * `'1:2'`, `columns` like `'A:B'` — or undefined when it repeats neither.
   * Reflects a pending `setPrintTitles`.
   */
  readonly printTitles: PrintTitles | undefined
  /**
   * Sets the rows and/or columns to repeat on each printed page, replacing any it
   * has. Pass `rows` as a row range like `'1:2'`, `columns` as a column range like
   * `'A:B'`; at least one is required, and an omitted axis is not repeated. Stored
   * as the built-in `_xlnm.Print_Titles` name scoped to the sheet.
   */
  setPrintTitles(titles: PrintTitles): void
  /** Removes the sheet's print titles, if it has any. */
  clearPrintTitles(): void
  /**
   * The worksheet protection in force, in the shape `protect()` takes, or
   * undefined when the sheet is not protected. Reflects a pending `protect()` or
   * `unprotect()` as well as what the file was read with.
   */
  readonly protection?: SheetProtection
  /**
   * Every cell the sheet stores. A cell that was cleared, that carries only
   * formatting, or that a comment sits on without a value of its own is still
   * stored, and arrives with a value of `kind: 'empty'`.
   *
   * Each call re-reads the sheet, so a call per cell is quadratic.
   */
  cells(): Iterable<Cell>
  /** Undefined when the sheet stores nothing at that reference. */
  cell(reference: string): Cell | undefined
  /**
   * Visible to `cells()` and `cell()` immediately, written by `toBytes()`.
   * A `numberFormat` is a format code such as `"$"#,##0.00`; without one the
   * cell keeps the formatting it already had.
   *
   * Throws `XlsxError` with code `unwritable-value` for a value the format
   * cannot hold, and records nothing when it does, so the rest of a batch of
   * edits still writes.
   */
  set(reference: string, value: CellInput, options?: SetOptions): void
  /**
   * Writes a block of values, `data[r][c]` landing `r` rows below and `c` columns
   * right of `topLeft`. Each cell goes through `set`, so the same refusals and
   * `options` apply — `options`, when given, styles every cell written. A row may
   * be shorter than the others; only the cells it holds are written.
   */
  setValues(
    topLeft: string,
    data: ReadonlyArray<ReadonlyArray<CellInput>>,
    options?: SetOptions,
  ): void
  /**
   * The values in a rectangular range, row by row, each cell as its `CellValue`;
   * a position the sheet stores nothing at is `{ kind: 'empty' }`. Reflects edits
   * made this session, like `cell`. A single reference reads a one-by-one block.
   */
  getValues(range: string): CellValue[][]
  /**
   * The sheet as delimited text, from `A1` to the furthest cell that holds
   * anything. A number, boolean, error or date prints as its value (a date in ISO
   * form); a formula prints its cached result, empty for one this library wrote
   * and never recalculated. A field that needs it is quoted, RFC 4180 style.
   */
  toCsv(options?: { readonly delimiter?: string }): string
  /**
   * Applies formatting to a cell without changing its value or formula, so a
   * formula cell can be restyled without losing its expression. A cell that is
   * not there yet is created empty with the formatting. Unlike `set`, this is
   * refused by nothing: restyling a shared-formula master or a merged cell is
   * safe, since only the cell format changes.
   */
  format(reference: string, options: SetOptions): void
  /**
   * Turns on worksheet protection, which is what makes a cell's `locked` and
   * `hidden` flags bite. Without `options` it matches Excel's Protect Sheet
   * default; `options` names the actions that stay permitted. Replaces any
   * protection the sheet already declared. Passwords are not written.
   */
  protect(options?: SheetProtection): void
  /** Removes worksheet protection, if the sheet had any. */
  unprotect(): void
  /** Whether the sheet shows gridlines. On unless a `showGridlines(false)` or the
   * file turned them off. */
  readonly gridlinesVisible: boolean
  /** Whether the sheet shows row and column headings. On by the same terms as
   * `gridlinesVisible`. */
  readonly headingsVisible: boolean
  /** The sheet's zoom as a whole percentage, or undefined at the default 100. */
  readonly zoomPercent: number | undefined
  /** The cell below-and-right of the frozen rows and columns, like `freeze` takes,
   * or undefined when the sheet is not frozen. */
  readonly frozenAt: string | undefined
  /** The tab colour as the 8-digit ARGB hex the file stores, or undefined for no
   * colour or one given as a theme or indexed reference this does not resolve. */
  readonly tabColorHex: string | undefined
  /** The sheet's auto-filter range, or undefined when it has none. */
  readonly autoFilterRange: string | undefined
  /**
   * The ranges the sheet merges, each as `A1:B2`, the file's own plus any added
   * this session with `merge`. Reflects the file as read; a shift from inserting
   * or deleting lines is applied on write, not here.
   */
  readonly mergedRanges: readonly string[]
  /**
   * Merges a rectangular range like `A1:B2`, joining any merges the sheet
   * already has. Excel shows only the top-left cell's value; the others keep
   * whatever they hold, since a merge does not clear them. Refuses a range that
   * is not two references either side of a colon.
   */
  merge(range: string): void
  /** Removes a merge over exactly `range`, by its `A1:B2` form. A range the sheet
   * does not merge is ignored; this does not split a merge it overlaps. */
  unmerge(range: string): void
  /** Sets the sheet's auto-filter over a range, replacing any it already has. */
  autoFilter(range: string): void
  /**
   * The tables the sheet carries, the file's own plus any added this session, each
   * with its range and column names.
   */
  readonly tables: readonly {
    readonly name: string
    readonly range: string
    readonly columns: readonly string[]
  }[]
  /**
   * Turns a range into a table with a banded style and a filter. The top row is the
   * header; its cells are set to the column names, taken from `options.columns`, the
   * header cells already there, or `Column1`, `Column2`… in order. Refuses a range
   * that is not two references, an empty or duplicate column name, a table name the
   * workbook already uses, or a range that overlaps a table the sheet already has.
   */
  addTable(
    range: string,
    options?: {
      readonly name?: string
      readonly columns?: readonly string[]
      readonly style?: string
    },
  ): void
  /**
   * Embeds a picture spanning `anchor` — a cell like `'B2'` or a range like
   * `'B2:E10'` the image fills, moving and staying sized with its cells. The type
   * (PNG, JPEG or GIF) is read from the image's own bytes; an unrecognised one is
   * refused with `unwritable-value`. The picture joins the sheet's drawing, one
   * being created if it has none, and is written by `toBytes()`.
   */
  addImage(image: Uint8Array, anchor: string): void
  /**
   * Adds a data validation over a cell or range. `{ list }` offers inline values as
   * a dropdown (no value may hold a comma, the inline list's separator) and
   * `{ listRange }` reads them from a range; `{ whole }`, `{ decimal }` and
   * `{ textLength }` constrain a number, a decimal or the text length against a
   * comparison, `{ date }` a date against `Date` bounds; `{ custom }` requires a
   * formula. Joins any the sheet already has. Written by `toBytes()`.
   */
  validate(range: string, rule: DataValidation): void
  /**
   * The data validations in force, each with the range it covers, the file's own
   * plus any added this session with `validate`. A rule of a kind this does not
   * model — a time — is left out.
   */
  readonly validations: readonly { readonly range: string; readonly rule: DataValidation }[]
  /** Removes every data validation on the sheet. A `validate` in the same session
   * still applies, so clear-then-validate leaves only the new rule. */
  clearValidations(): void
  /**
   * Adds a conditional format over a cell or range. `{ colorScale }` grades cells
   * between two colours, or three with a `mid`; `{ cellIs }` fills cells matching a
   * comparison, `{ expression }` cells matching a formula, `{ duplicates }` and
   * `{ unique }` the repeated or one-off values, `{ top }` and `{ bottom }` the
   * highest or lowest ranked; `{ dataBar }` draws a bar. The rule outranks any the
   * sheet already has, and is written by `toBytes()`.
   */
  conditionalFormat(range: string, rule: ConditionalFormat): void
  /**
   * The conditional formats in force, each with the range it covers, the file's
   * own plus any added this session. A rule of a kind this does not model — a
   * colour scale with a theme-coloured stop, a `cellIs` whose highlight is not a
   * plain colour — is left out.
   */
  readonly conditionalFormats: readonly {
    readonly range: string
    readonly rule: ConditionalFormat
  }[]
  /** Removes every conditional format on the sheet. A `conditionalFormat` in the
   * same session still applies, so clear-then-add leaves only the new rule. */
  clearConditionalFormats(): void
  /**
   * Attaches a comment to a cell, written by `toBytes()`. A sheet that already
   * has comments is added to in place, its existing rich text kept. Refused with
   * `unsupported-edit` only on a cell that already carries one, which this does
   * not replace.
   */
  comment(reference: string, text: string): void
  /** Removes a cell's comment, its note box included. A cell without one is
   * ignored; the comments part is left in place even when it empties. */
  removeComment(reference: string): void
  /**
   * Sets the sheet tab's colour, into `sheetPr`, replacing any it already has.
   * The colour is a 6- or 8-digit hex string; a 6-digit one gains an opaque
   * alpha. Refuses anything else. Written by `toBytes()`.
   */
  tabColor(color: string): void
  /** Shows or hides the sheet's gridlines, as `showGridLines` on its view. */
  showGridlines(visible: boolean): void
  /** Shows or hides the sheet's row and column headings, as `showRowColHeaders`. */
  showHeadings(visible: boolean): void
  /** Sets the sheet's zoom as a whole percentage. Refuses one outside 10 to 400. */
  zoom(percent: number): void
  /**
   * Groups rows `from` to `to` at an outline level (1 to 7, default 1), so a
   * reader shows a collapsible band. Deeper levels nest inside shallower ones.
   * Refuses a backwards range or a level out of bounds.
   */
  groupRows(from: number, to: number, level?: number): void
  /**
   * Groups columns `from` to `to` — letters like `B` and `D` — at an outline
   * level, on the same terms as `groupRows`.
   */
  groupColumns(from: string, to: string, level?: number): void
  /** The outline level a row sits at, 1 or more, or 0 when it is not grouped. In
   * the file or by a `groupRows` this session. The row is one-based. */
  rowGroupLevel(row: number): number
  /** The outline level a column sits at, or 0 when it is not grouped, by the same
   * terms as `rowGroupLevel`. The column is a letter like `A`. */
  columnGroupLevel(column: string): number
  /**
   * Freezes the rows above and the columns left of `cell`, so they stay in view
   * when the sheet is scrolled. `freeze('B2')` freezes row 1 and column A.
   */
  freeze(cell: string): void
  /** The page setup for printing — orientation and scale — the file's plus any set
   * this session. */
  readonly pageSetup: PageSetup
  /** Sets the page orientation and print scale, merging onto what the sheet has and
   * keeping any other `pageSetup` attribute. Refuses a scale outside 10 to 400. */
  setPageSetup(setup: PageSetup): void
  /** The print margins in inches, the file's own plus any set this session. */
  readonly pageMargins: PageMargins
  /** Sets print margins, merging onto the ones the sheet has, so a partial edit
   * still leaves all six. Refuses a margin that is not a finite number at least zero. */
  setPageMargins(margins: PageMargins): void
  /** Whether the sheet prints its gridlines and its row and column headings, each
   * off unless the file or a `setPrintOptions` this session turned it on. */
  readonly printOptions: { readonly gridlines: boolean; readonly headings: boolean }
  /** Sets whether gridlines and headings print, merging onto what the sheet has, so
   * setting one leaves the other. */
  setPrintOptions(options: PrintOptions): void
  /** The printed header and footer, each split into left, centre and right, the
   * file's plus any set this session. */
  readonly headerFooter: HeaderFooter
  /**
   * Sets the printed header and footer. Each section string is Excel field-code
   * text — `&P` the page number, `&N` the page count, `&D` the date, `&&` a literal
   * ampersand. The `header` or `footer` passed is replaced whole; omit one to leave
   * it, and any even- or first-page variant the file carries is kept.
   */
  setHeaderFooter(headerFooter: HeaderFooter): void
  /** The manual page breaks, the file's plus any added this session: the one-based
   * rows and the column letters that begin a new page. */
  readonly pageBreaks: PageBreaks
  /**
   * Adds a manual page break above `row`, so `row` begins a new page when printed.
   * The row is one-based. Refuses a row below 2 (nothing sits above row 1) or past
   * the sheet's last row.
   */
  addRowPageBreak(row: number): void
  /**
   * Adds a manual page break to the left of `column`, so `column` begins a new page
   * when printed. The column is a letter like `D`. Refuses column `A` (nothing sits
   * left of it) or a column past the sheet's last.
   */
  addColumnPageBreak(column: string): void
  /**
   * A row's height in points, the file's or one set this session, or undefined
   * when the row carries no height of its own. The row is one-based.
   */
  rowHeight(row: number): number | undefined
  /**
   * A column's width in the units Excel shows, the file's or one set this
   * session, or undefined when no `<col>` covering it carries a width. The column
   * is a letter like `A`.
   */
  columnWidth(column: string): number | undefined
  /**
   * Sets a row's height in points, marking it a custom height so a reader keeps
   * it. Refuses a row number below 1 or a height that is not a finite number at
   * least zero.
   */
  setRowHeight(row: number, height: number): void
  /**
   * Sets a column's width, in the units Excel shows, splitting a `cols` range
   * that spans more than the column so the rest keeps its own width. The column
   * is a letter like `A`. Refuses a width that is not a finite number at least
   * zero.
   */
  setColumnWidth(column: string, width: number): void
  /** Whether the row is hidden, in the file or by a `hideRow` this session. */
  isRowHidden(row: number): boolean
  /** Whether the column is hidden, in the file or by a `hideColumn` this session.
   * The column is a letter like `A`. */
  isColumnHidden(column: string): boolean
  /** Hides a row, keeping any height it also has. The row is one-based. */
  hideRow(row: number): void
  /** Hides a column, keeping any width it also has. The column is a letter. */
  hideColumn(column: string): void
  /**
   * Inserts `count` blank rows before row `before`, pushing the rows at and below
   * it down. References that point into the moved rows — formulas anywhere in the
   * workbook, merges, the dimension, filters, conditional formats and defined
   * names — move with them. Cell edits made in this session land first, so they
   * ride along too. Refused when the sheet carries a table, drawing, pivot table
   * or comment, whose pinned positions this does not yet adjust, or when the
   * shift would push a row off the sheet.
   */
  insertRows(before: number, count?: number): void
  /**
   * Inserts `count` blank columns before column `before` — a letter like `C` —
   * pushing the columns at and to the right of it over. References that point into
   * the moved columns move with them, on the same terms as `insertRows`, a pinned
   * part and an off-the-sheet shift included.
   */
  insertColumns(before: string, count?: number): void
  /**
   * Deletes `count` rows from row `from`, pulling the rows below up over them. A
   * reference into a deleted row becomes #REF! where a formula named the cell, and
   * shrinks where a range only overlapped. Refused when a whole merged range,
   * filter, format or shared formula sat inside the deletion, none of which can
   * survive as #REF!, or when the sheet carries a table, drawing, pivot table or
   * comment.
   */
  deleteRows(from: number, count?: number): void
  /**
   * Deletes `count` columns from column `from` — a letter like `C` — pulling the
   * columns to their right in over them, on the same terms as `deleteRows`.
   */
  deleteColumns(from: string, count?: number): void
  /**
   * Links a cell to a URL or a place in the workbook, replacing any link the cell
   * already has. An external URL is written through a worksheet relationship; a
   * `location` is written inline. Visible in the file after `toBytes()`.
   */
  link(reference: string, target: Hyperlink): void
  /** Removes the link anchored at a cell. A cell without one is ignored; the
   * external relationship a URL used is left as a harmless dangling entry. */
  unlink(reference: string): void
}

/** A comparison keyed by operator: a range for `between`/`notBetween`, a single
 * bound otherwise. The bound type varies — a number for a numeric rule, a `Date`
 * for a date rule. */
export type Constraint<T> =
  | { readonly between: readonly [T, T] }
  | { readonly notBetween: readonly [T, T] }
  | { readonly equal: T }
  | { readonly notEqual: T }
  | { readonly greaterThan: T }
  | { readonly lessThan: T }
  | { readonly greaterThanOrEqual: T }
  | { readonly lessThanOrEqual: T }

/** A comparison for a `whole`, `decimal` or `textLength` rule. Each bound is a
 * finite number. */
export type NumberConstraint = Constraint<number>

/** A comparison for a `date` rule. Each bound is a `Date`. */
export type DateConstraint = Constraint<Date>

/**
 * A data-validation rule, keyed by kind: a `list` dropdown, or a `whole` or
 * `decimal` numeric comparison. The union is open to more kinds without a break.
 */
export type DataValidation = { readonly allowBlank?: boolean } & (
  | {
      /** The values a cell may take, as a dropdown. At least one, none with a
       * comma — the inline list separates values with commas. */
      readonly list: readonly string[]
    }
  /** A dropdown reading its values from a range, like `Sheet1!$A$1:$A$10`. */
  | { readonly listRange: string }
  | { readonly whole: NumberConstraint }
  | { readonly decimal: NumberConstraint }
  /** Constrains the cell's text length, the same comparisons as `whole`. */
  | { readonly textLength: NumberConstraint }
  /** Constrains the cell's date against `Date` bounds. */
  | { readonly date: DateConstraint }
  /** A formula the cell must satisfy, written verbatim (`ISNUMBER(A1)`). */
  | { readonly custom: string }
)

/** A colour scale graded across a range: two stops, or three with a midpoint. */
export interface ColorScale {
  /** Hex for the lowest value. */
  readonly min: string
  /** Hex for the midpoint (the 50th percentile), making it a three-colour scale. */
  readonly mid?: string
  /** Hex for the highest value. */
  readonly max: string
}

/** Highlights a cell whose value meets `when` by filling it with `fill` (hex). */
export interface CellValueRule {
  readonly when: NumberConstraint
  readonly fill: string
}

/** Fills a cell when a formula holds. `formula` is written verbatim; `fill` is hex. */
export interface FormulaRule {
  readonly formula: string
  readonly fill: string
}

/** Fills the cells a rule matches. `fill` is hex. */
export interface FillRule {
  readonly fill: string
}

/** Fills the top or bottom `count` cells of a range. With `percent`, `count` is a
 * percentage rather than a number of cells. `fill` is hex. */
export interface RankRule {
  readonly count: number
  readonly fill: string
  readonly percent?: boolean
}

/** A bar drawn in each cell, its length scaled between the range's min and max. */
export interface DataBar {
  /** The bar's colour, hex. */
  readonly color: string
}

/**
 * A conditional-format rule, keyed by kind: a `colorScale` graded across a range,
 * a `cellIs` value comparison that fills the cells it matches, or a `dataBar`
 * scaled across the range. The union is open to more kinds without a break.
 */
export type ConditionalFormat =
  | { readonly colorScale: ColorScale }
  | { readonly cellIs: CellValueRule }
  | { readonly dataBar: DataBar }
  /** Fills a cell whose value satisfies a formula, like `$B1>0`. */
  | { readonly expression: FormulaRule }
  /** Fills the cells whose value is duplicated within the range. */
  | { readonly duplicates: FillRule }
  /** Fills the cells whose value is unique within the range. */
  | { readonly unique: FillRule }
  /** Fills the highest-valued cells of the range. */
  | { readonly top: RankRule }
  /** Fills the lowest-valued cells of the range. */
  | { readonly bottom: RankRule }

export interface SetOptions {
  /** A number format code, applied to the cell being written. */
  readonly numberFormat?: string
  /** Font to apply, merged onto the font the cell already carries. */
  readonly font?: FontFormat
  /** The cell's background fill: solid, a pattern, or a gradient. */
  readonly fill?: FillFormat
  /** Borders to apply, merged onto the sides the cell already has. */
  readonly border?: BorderFormat
  /** Alignment to apply, merged onto the alignment the cell already has. */
  readonly alignment?: Alignment
  /** Protection to apply, merged onto the protection the cell already has. */
  readonly protection?: CellProtection
}

export interface Workbook {
  readonly sheets: readonly Worksheet[]
  /** Undefined when no sheet has that name. Names are compared exactly. */
  sheet(name: string): Worksheet | undefined
  /**
   * Adds an empty worksheet and returns it, ready to fill. The name must be one
   * Excel accepts — up to 31 characters, none of `: \ / ? * [ ]`, and not one a
   * sheet already uses. Written into the workbook by `toBytes()`.
   */
  addSheet(name: string): Worksheet
  /**
   * Duplicates a sheet under a new name and returns the copy, ready to edit. The
   * cells, formatting, formulas, merges, validations, conditional formats, widths
   * and page setup all come across, as do comments and printer settings. A sheet
   * that carries a table, a drawing or a pivot table is refused with
   * `unsupported-edit`, since those need names, ids or media reworked to stay
   * valid. `newName` follows the same rules as `addSheet`.
   */
  copySheet(sourceName: string, newName: string): Worksheet
  /**
   * The workbook's global named ranges, each mapped to what it refers to (a
   * formula like `Sheet1!$A$1:$B$2`). Reflects a pending `defineName`.
   */
  readonly definedNames: ReadonlyMap<string, string>
  /**
   * Defines a global named range, replacing one of the same name. The name
   * follows Excel's rules — a letter, underscore or backslash then letters,
   * digits, periods or underscores, no spaces. `refersTo` is a formula.
   */
  defineName(name: string, refersTo: string): void
  /** Removes a global named range. A name the workbook does not have is ignored. */
  removeDefinedName(name: string): void
  /**
   * The document's core properties (title, creator, dates and the rest), the
   * file's own plus any set this session. Empty fields are absent, not blank.
   */
  readonly properties: DocumentProperties
  /**
   * Sets core properties, each replacing the one the file had and leaving the
   * others — and any property this does not model — as they were. Written by
   * `toBytes()`, into a fresh `docProps/core.xml` when the file has none.
   */
  setProperties(properties: DocumentProperties): void
  /** Which year serials count from. A 1904 workbook is 1462 days behind. */
  readonly epoch: 1900 | 1904
  /**
   * Resolves a stored colour reference — the `{ theme, tint }` or `{ indexed }` a
   * cell's font, fill or border can carry — to the 8-digit ARGB hex it displays
   * as. A plain hex passes through. Returns undefined for a colour with no fixed
   * value: a system indexed colour, or a theme slot this workbook's theme does
   * not define. A tinted theme colour is resolved in the colour space Excel uses
   * and matches its shown value to within one unit per channel.
   */
  resolveColor(color: Color): string | undefined
  /** Parts that were never interpreted are written exactly as they were read. */
  toBytes(): Uint8Array
}

/** Options for `createWorkbookFromCsv`. */
export interface CsvReadOptions {
  /** The name the one sheet is given; `Sheet1` by default. */
  readonly sheetName?: string
  /** The field separator; a comma by default. Pass `'\t'` for TSV. */
  readonly delimiter?: string
  /** When true, a field that is a plain finite number is written as one rather
   * than as text. Off by default, since coercing loses a leading zero and turns
   * a code like `007` into `7`. */
  readonly parseNumbers?: boolean
}
