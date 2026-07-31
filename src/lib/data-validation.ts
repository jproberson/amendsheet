import { comparisonToConstraint, mapConstraint, numberComparison } from './constraint.js'
import { dateToSerial, serialToDate } from './date.js'
import { XlsxError } from './errors.js'
import type { DataValidationSpec, SheetLocation } from './patch.js'
import type { DataValidation } from './public-types.js'

function buildValidationSpec(
  rule: DataValidation,
  sqref: string,
  at: SheetLocation,
  date1904: boolean,
): DataValidationSpec {
  const allowBlank = rule.allowBlank ?? true
  if ('list' in rule) {
    if (rule.list.length === 0) {
      throw new XlsxError('unwritable-value', 'A list validation needs at least one value', {
        ...at,
        reference: sqref,
      })
    }
    for (const value of rule.list) {
      if (value.includes(',')) {
        throw new XlsxError(
          'unwritable-value',
          `List value "${value}" holds a comma, which an inline list reads as the next value`,
          { ...at, reference: sqref },
        )
      }
    }
    return { type: 'list', sqref, allowBlank, formula1: `"${rule.list.join(',')}"` }
  }
  if ('listRange' in rule) {
    return { type: 'list', sqref, allowBlank, formula1: rule.listRange }
  }
  if ('custom' in rule) {
    return { type: 'custom', sqref, allowBlank, formula1: rule.custom }
  }

  // A date rule compares against serials, so its Date bounds become the same
  // number rules do; the type marks it a date so a reader turns them back.
  const type =
    'whole' in rule
      ? 'whole'
      : 'decimal' in rule
        ? 'decimal'
        : 'textLength' in rule
          ? 'textLength'
          : 'date'
  const constraint =
    'whole' in rule
      ? rule.whole
      : 'decimal' in rule
        ? rule.decimal
        : 'textLength' in rule
          ? rule.textLength
          : mapConstraint(rule.date, (date) => dateToSerial(date, date1904))
  const comparison = numberComparison(constraint)
  const bounds =
    comparison.formula2 === undefined
      ? [comparison.formula1]
      : [comparison.formula1, comparison.formula2]
  for (const bound of bounds) {
    if (!Number.isFinite(bound)) {
      throw new XlsxError('unwritable-value', `Validation bound ${bound} is not a finite number`, {
        ...at,
        reference: sqref,
      })
    }
  }
  return {
    type,
    sqref,
    allowBlank,
    operator: comparison.operator,
    formula1: String(comparison.formula1),
    formula2: comparison.formula2 === undefined ? undefined : String(comparison.formula2),
  }
}

/** The list a validation offers, or undefined when it names a range not an
 * inline set — only the inline form maps back to string values. */
function listFromFormula(formula1: string): readonly string[] | undefined {
  if (formula1.length < 2 || !formula1.startsWith('"') || !formula1.endsWith('"')) return undefined
  const inner = formula1.slice(1, -1)
  return inner === '' ? [] : inner.split(',')
}

/** A stored validation spec back into the public rule, or undefined for a type
 * this does not model (a time). */
function validationFromSpec(
  spec: DataValidationSpec,
  date1904: boolean,
): DataValidation | undefined {
  if (spec.type === 'list') {
    const list = listFromFormula(spec.formula1)
    return list === undefined
      ? { allowBlank: spec.allowBlank, listRange: spec.formula1 }
      : { allowBlank: spec.allowBlank, list }
  }
  if (spec.type === 'custom') {
    return { allowBlank: spec.allowBlank, custom: spec.formula1 }
  }
  const constraint = comparisonToConstraint(spec.operator, spec.formula1, spec.formula2)
  if (constraint === undefined) return undefined
  if (spec.type === 'whole') return { allowBlank: spec.allowBlank, whole: constraint }
  if (spec.type === 'decimal') return { allowBlank: spec.allowBlank, decimal: constraint }
  if (spec.type === 'textLength') return { allowBlank: spec.allowBlank, textLength: constraint }
  if (spec.type === 'date') {
    return {
      allowBlank: spec.allowBlank,
      date: mapConstraint(constraint, (serial) => serialToDate(serial, date1904)),
    }
  }
  return undefined
}

/**
 * The pending data validations for one workbook: rules added this session, per
 * sheet path, and the sheets whose file rules were cleared. Building a rule into
 * a spec, mapping a spec back to a rule, and merging file rules with pending ones
 * all live here; `document.ts` supplies the sheet's own reference plumbing and
 * the file rules it reads.
 */
export interface ValidationStore {
  /** Pending specs per sheet path. Exposed for the write pass's sheet-rewrite
   * set and the `SheetEdits` it hands `patchSheet`. */
  readonly pending: ReadonlyMap<string, DataValidationSpec[]>
  /** Sheets whose file validations were dropped this session. */
  readonly cleared: ReadonlySet<string>
  add(path: string, sqref: string, at: SheetLocation, rule: DataValidation): void
  clear(path: string): void
  /** File rules (already read by the caller) plus pending ones, each mapped to
   * the public form, minus any kind this does not model. */
  applied(
    path: string,
    fileSpecs: readonly DataValidationSpec[],
  ): { range: string; rule: DataValidation }[]
}

export function createValidationStore(date1904: boolean): ValidationStore {
  const pending = new Map<string, DataValidationSpec[]>()
  const cleared = new Set<string>()
  return {
    pending,
    cleared,
    add(path, sqref, at, rule) {
      const spec = buildValidationSpec(rule, sqref, at, date1904)
      const specs = pending.get(path) ?? []
      specs.push(spec)
      pending.set(path, specs)
    },
    clear(path) {
      pending.delete(path)
      cleared.add(path)
    },
    applied(path, fileSpecs) {
      const fromFile = cleared.has(path) ? [] : fileSpecs
      const out: { range: string; rule: DataValidation }[] = []
      for (const spec of [...fromFile, ...(pending.get(path) ?? [])]) {
        const rule = validationFromSpec(spec, date1904)
        if (rule !== undefined) out.push({ range: spec.sqref, rule })
      }
      return out
    },
  }
}
