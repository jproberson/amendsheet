import { comparisonToConstraint } from './constraint.js'
import type { ConditionalFormatSpec } from './patch.js'
import type { ConditionalFormat, RankRule } from './public-types.js'
import { readDxfFill } from './styles-writer.js'

/** A stored conditional-format spec back into the public rule, or undefined for
 * one this does not model: a colour scale short of two rgb stops, a highlight rule
 * whose colour or comparison cannot be recovered. `stylesXml` resolves the
 * highlight a dxf-backed rule names by its index. */
function conditionalFormatFromSpec(
  spec: ConditionalFormatSpec,
  stylesXml: string | undefined,
): ConditionalFormat | undefined {
  if (spec.kind === 'dataBar') return { dataBar: { color: spec.color } }
  if (spec.kind === 'colorScale') {
    const [min, second, third] = spec.colors
    if (min === undefined || second === undefined) return undefined
    return third === undefined
      ? { colorScale: { min, max: second } }
      : { colorScale: { min, mid: second, max: third } }
  }
  const fill = stylesXml === undefined ? undefined : readDxfFill(stylesXml, spec.dxfId)
  if (fill === undefined) return undefined
  if (spec.kind === 'expression') return { expression: { formula: spec.formula, fill } }
  if (spec.kind === 'cellIs') {
    const when = comparisonToConstraint(spec.operator, spec.formulas[0] ?? '', spec.formulas[1])
    return when === undefined ? undefined : { cellIs: { when, fill } }
  }
  if (spec.kind === 'top10') {
    const rank: RankRule = spec.percent
      ? { count: spec.rank, fill, percent: true }
      : { count: spec.rank, fill }
    return spec.bottom ? { bottom: rank } : { top: rank }
  }
  // Only duplicateValues and uniqueValues remain.
  return spec.kind === 'duplicateValues' ? { duplicates: { fill } } : { unique: { fill } }
}

/**
 * The pending conditional formats for one workbook: specs added this session,
 * per sheet path, and the sheets whose file rules were cleared. Building a rule
 * into a spec stays in `document.ts` because it allocates a dxf highlight in the
 * styles table; mapping a spec back to a rule and merging file rules with pending
 * ones live here, mirroring the data-validation store.
 */
export interface ConditionalFormatStore {
  /** Pending specs per sheet path, for the write pass's sheet-rewrite set and
   * the `SheetEdits` it hands `patchSheet`. */
  readonly pending: ReadonlyMap<string, ConditionalFormatSpec[]>
  /** Sheets whose file conditional formats were dropped this session. */
  readonly cleared: ReadonlySet<string>
  add(path: string, spec: ConditionalFormatSpec): void
  clear(path: string): void
  /** File specs (already read by the caller) plus pending ones, each mapped to
   * the public form, minus any kind this does not model. `stylesXml` resolves a
   * dxf-backed highlight. */
  applied(
    path: string,
    fileSpecs: readonly ConditionalFormatSpec[],
    stylesXml: string | undefined,
  ): { range: string; rule: ConditionalFormat }[]
}

export function createConditionalFormatStore(): ConditionalFormatStore {
  const pending = new Map<string, ConditionalFormatSpec[]>()
  const cleared = new Set<string>()
  return {
    pending,
    cleared,
    add(path, spec) {
      const specs = pending.get(path) ?? []
      specs.push(spec)
      pending.set(path, specs)
    },
    clear(path) {
      pending.delete(path)
      cleared.add(path)
    },
    applied(path, fileSpecs, stylesXml) {
      const fromFile = cleared.has(path) ? [] : fileSpecs
      const out: { range: string; rule: ConditionalFormat }[] = []
      for (const spec of [...fromFile, ...(pending.get(path) ?? [])]) {
        const rule = conditionalFormatFromSpec(spec, stylesXml)
        if (rule !== undefined) out.push({ range: spec.sqref, rule })
      }
      return out
    },
  }
}
