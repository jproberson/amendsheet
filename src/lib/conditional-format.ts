import type { SheetLocation } from './cell-input.js'
import { comparisonToConstraint, numberComparison } from './constraint.js'
import { XlsxError } from './errors.js'
import type { ConditionalFormatSpec } from './patch.js'
import type { ConditionalFormat, RankRule } from './public-types.js'
import { ensureDxf, normalizeColor, readDxfFill } from './styles-writer.js'

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
 * Works out the conditional-format specs a call to `conditionalFormat` should stage
 * from one public rule. A colour scale or data bar maps straight across; every
 * highlight rule (cellIs, expression, duplicates, unique, top/bottom) needs a dxf
 * in the styles table, so this allocates one through `ensureDxf`, threading the
 * styles xml and collecting the colours the caller folds onto the serialized styles.
 * Refusals — no style table to hold a highlight, a non-finite comparison bound, a
 * rank that is not a positive whole number — are raised here.
 */
export function planConditionalFormat(
  rule: ConditionalFormat,
  sqref: string,
  at: SheetLocation,
  dxfStyles: string | undefined,
): { specs: ConditionalFormatSpec[]; dxfStyles: string | undefined; dxfColors: string[] } {
  if ('colorScale' in rule) {
    const { min, mid, max } = rule.colorScale
    const colors =
      mid === undefined
        ? [normalizeColor(min, at), normalizeColor(max, at)]
        : [normalizeColor(min, at), normalizeColor(mid, at), normalizeColor(max, at)]
    return { specs: [{ kind: 'colorScale', sqref, colors }], dxfStyles, dxfColors: [] }
  }
  if ('dataBar' in rule) {
    const color = normalizeColor(rule.dataBar.color, at)
    return { specs: [{ kind: 'dataBar', sqref, color }], dxfStyles, dxfColors: [] }
  }

  // The rest fill matching cells with a highlight held in a dxf.
  if (dxfStyles === undefined) {
    throw new XlsxError(
      'missing-part',
      `Cannot fill ${sqref}: the package has no style table to hold the format`,
      { ...at, part: 'xl/styles.xml', reference: sqref },
    )
  }
  const highlight =
    'cellIs' in rule
      ? rule.cellIs.fill
      : 'expression' in rule
        ? rule.expression.fill
        : 'duplicates' in rule
          ? rule.duplicates.fill
          : 'unique' in rule
            ? rule.unique.fill
            : 'top' in rule
              ? rule.top.fill
              : rule.bottom.fill
  const color = normalizeColor(highlight, at)
  const dxf = ensureDxf(dxfStyles, color)
  const dxfId = dxf.index

  let spec: ConditionalFormatSpec
  if ('top' in rule || 'bottom' in rule) {
    const rank = 'top' in rule ? rule.top : rule.bottom
    if (!Number.isInteger(rank.count) || rank.count < 1) {
      throw new XlsxError('unwritable-value', `Rank ${rank.count} is not a positive whole number`, {
        ...at,
        reference: sqref,
      })
    }
    spec = {
      kind: 'top10',
      sqref,
      rank: rank.count,
      bottom: 'bottom' in rule,
      percent: rank.percent ?? false,
      dxfId,
    }
  } else if ('cellIs' in rule) {
    const comparison = numberComparison(rule.cellIs.when)
    const formulas =
      comparison.formula2 === undefined
        ? [comparison.formula1]
        : [comparison.formula1, comparison.formula2]
    for (const bound of formulas) {
      if (!Number.isFinite(bound)) {
        throw new XlsxError(
          'unwritable-value',
          `Conditional-format bound ${bound} is not a finite number`,
          { ...at, reference: sqref },
        )
      }
    }
    spec = {
      kind: 'cellIs',
      sqref,
      operator: comparison.operator,
      formulas: formulas.map(String),
      dxfId,
    }
  } else if ('expression' in rule) {
    spec = { kind: 'expression', sqref, formula: rule.expression.formula, dxfId }
  } else if ('duplicates' in rule) {
    spec = { kind: 'duplicateValues', sqref, dxfId }
  } else {
    spec = { kind: 'uniqueValues', sqref, dxfId }
  }
  return { specs: [spec], dxfStyles: dxf.xml, dxfColors: [color] }
}

/**
 * The pending conditional formats for one workbook: specs added this session,
 * per sheet path, and the sheets whose file rules were cleared. `planConditionalFormat`
 * (above) builds a rule into specs; this stores them, and `applied` maps a spec back
 * to a rule and merges file rules with pending ones, mirroring the data-validation store.
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
