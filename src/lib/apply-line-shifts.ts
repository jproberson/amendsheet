import {
  COMMENTS_RELATIONSHIP,
  VML_DRAWING_RELATIONSHIP,
  shiftComments,
  shiftNoteShapes,
} from './comments.js'
import type { ContainerDraft } from './container-draft.js'
import type { Container } from './container.js'
import { shiftDrawing } from './drawings.js'
import type { SheetLocation } from './cell-input.js'
import { patchSheet } from './patch.js'
import { shiftPivotCacheSource, shiftPivotLocation } from './pivots.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { type ShiftSpec, shiftDefinedNames } from './shift.js'
import { shiftForeignFormulas, shiftSheet } from './shift-sheet.js'
import { insertTableColumns, shiftTables } from './tables.js'
import { relationshipsPathFor } from './workbook-parts.js'

export interface LineOp {
  readonly path: string
  readonly spec: ShiftSpec
}

/** What `applyLineShifts` needs about the workbook to shift its references: the
 * sheets and any added this session, which are removed, the line ops in call
 * order, the file's and this session's defined names, and the date epoch. */
export interface ShiftInputs {
  readonly sheets: readonly { readonly path: string; readonly name: string }[]
  readonly addedSheetPaths: Iterable<string>
  readonly removed: ReadonlySet<string>
  readonly lineOps: readonly LineOp[]
  readonly fileNames: ReadonlyMap<string, string>
  readonly pendingNames: ReadonlyMap<string, string>
  readonly date1904: boolean
}

/**
 * Inserting or deleting a line moves references across the whole workbook. The
 * per-sheet patch has already landed this session's edits in the old grid, so each
 * sheet's current text is shifted here: the edited sheet's own rows and cells, and
 * every other sheet's formulas — and charts, pivots, tables, comments and drawings —
 * that point into it. Returns the workbook's defined names shifted the same way,
 * for the caller to write; when nothing shifts it returns `pendingNames` unchanged.
 */
export function applyLineShifts(
  draft: ContainerDraft,
  container: Container,
  changes: Map<string, Uint8Array | null>,
  input: ShiftInputs,
): ReadonlyMap<string, string> {
  const { sheets, addedSheetPaths, removed, lineOps, fileNames, pendingNames, date1904 } = input
  if (lineOps.length === 0) return pendingNames

  const encoder = new TextEncoder()
  const sheetPaths = [...sheets.map((sheet) => sheet.path), ...addedSheetPaths]
  for (const path of sheetPaths) {
    if (removed.has(path)) continue
    let xml = draft.text(path)
    if (xml === undefined) continue
    const before = xml
    for (const op of lineOps) {
      xml =
        op.path === path
          ? shiftSheet(xml, op.spec)
          : shiftForeignFormulas(xml, { ...op.spec, onCurrentSheet: false })
    }
    if (xml !== before) changes.set(path, encoder.encode(xml))

    // The sheet's cells have moved; each table part it owns carries a range that
    // must move with them. A grow earlier this session may already sit in `changes`,
    // so shift from there when present.
    const sheetOriginal = container.parts.get(path)
    if (sheetOriginal !== undefined) {
      const latestTable = (tablePath: string) => changes.get(tablePath) ?? undefined
      for (const op of lineOps) {
        if (op.path !== path) continue
        for (const extension of shiftTables(latestTable, sheetOriginal, path, container, op.spec)) {
          changes.set(extension.path, encoder.encode(extension.xml))
        }
        // A column inserted inside a table gains a fresh column entry, and its
        // header cell is authored on the shifted sheet so no named column sits over
        // a blank header. The header is written as an inline string.
        for (const insert of insertTableColumns(
          latestTable,
          sheetOriginal,
          path,
          container,
          op.spec,
        )) {
          changes.set(insert.path, encoder.encode(insert.xml))
          const sheetNow = draft.text(path)
          if (insert.headers.size > 0 && sheetNow !== undefined) {
            const at: SheetLocation = {
              sheet: sheets.find((sheet) => sheet.path === path)?.name,
              part: path,
            }
            changes.set(
              path,
              patchSheet(
                encoder.encode(sheetNow),
                insert.headers,
                date1904,
                undefined,
                undefined,
                at,
              ),
            )
          }
        }
      }
    }

    // A comment pins its cell in the comments part and its box in the legacy
    // drawing; both move with the rows or columns the edit shifts under them, and a
    // note whose cell a deletion removed is dropped from each.
    const commentsPath = draft.relationshipTarget(path, COMMENTS_RELATIONSHIP)
    const vmlPath = draft.relationshipTarget(path, VML_DRAWING_RELATIONSHIP)
    const sheetRels = draft.text(relationshipsPathFor(path))
    for (const op of lineOps) {
      if (op.path !== path) continue
      const commentsText = commentsPath === undefined ? undefined : draft.text(commentsPath)
      if (commentsPath !== undefined && commentsText !== undefined) {
        const shifted = shiftComments(commentsText, op.spec)
        if (shifted !== commentsText) changes.set(commentsPath, encoder.encode(shifted))
      }
      const vmlText = vmlPath === undefined ? undefined : draft.text(vmlPath)
      if (vmlPath !== undefined && vmlText !== undefined) {
        const shifted = shiftNoteShapes(vmlText, op.spec)
        if (shifted !== vmlText) changes.set(vmlPath, encoder.encode(shifted))
      }
      // A drawing (a diagram-bearing one is refused at the call) moves its cell
      // anchors with the edit; a picture fully inside a deletion goes. A pivot table
      // on this sheet moves its location, the range where it is drawn.
      if (sheetRels !== undefined) {
        for (const relationship of readRelationships(sheetRels, path).values()) {
          if (relationship.external) continue
          const target = resolveTarget(path, relationship.target)
          const partXml = draft.text(target)
          if (partXml === undefined) continue
          if (relationship.type.endsWith('relationships/drawing')) {
            const shifted = shiftDrawing(partXml, op.spec)
            if (shifted !== partXml) changes.set(target, encoder.encode(shifted))
          } else if (relationship.type.endsWith('relationships/pivotTable')) {
            const shifted = shiftPivotLocation(partXml, op.spec)
            if (shifted !== partXml) changes.set(target, encoder.encode(shifted))
          }
        }
      }
    }
  }

  // A chart's series, category and title references name a sheet and shift like any
  // foreign formula when that sheet's rows or columns move — wherever the chart part
  // lives, so a chart plotting another sheet is caught too. The drawing anchor moved
  // above; this moves the data the chart plots.
  for (const chartPath of container.parts.keys()) {
    if (!/^xl\/charts\/chart\d+\.xml$/.test(chartPath)) continue
    let xml = draft.text(chartPath)
    if (xml === undefined) continue
    const before = xml
    for (const op of lineOps) {
      xml = shiftForeignFormulas(xml, { ...op.spec, onCurrentSheet: false })
    }
    if (xml !== before) changes.set(chartPath, encoder.encode(xml))
  }

  // A pivot cache's source names its sheet, so its range shifts when that sheet is
  // edited, wherever the pivot that reads it sits. The location moved above.
  for (const cachePath of container.parts.keys()) {
    if (!/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(cachePath)) continue
    let xml = draft.text(cachePath)
    if (xml === undefined) continue
    const before = xml
    for (const op of lineOps) {
      xml = shiftPivotCacheSource(xml, op.spec)
    }
    if (xml !== before) changes.set(cachePath, encoder.encode(xml))
  }

  return shiftDefinedNames(
    fileNames,
    pendingNames,
    lineOps.map((op) => op.spec),
  )
}
