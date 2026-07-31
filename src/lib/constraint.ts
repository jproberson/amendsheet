import type { Constraint, NumberConstraint } from './public-types.js'

/** Maps a constraint's bounds through `f`, preserving its operator — a date
 * constraint into the serial one written, and back again on read. */
export function mapConstraint<A, B>(constraint: Constraint<A>, f: (bound: A) => B): Constraint<B> {
  if ('between' in constraint)
    return { between: [f(constraint.between[0]), f(constraint.between[1])] }
  if ('notBetween' in constraint)
    return { notBetween: [f(constraint.notBetween[0]), f(constraint.notBetween[1])] }
  if ('equal' in constraint) return { equal: f(constraint.equal) }
  if ('notEqual' in constraint) return { notEqual: f(constraint.notEqual) }
  if ('greaterThan' in constraint) return { greaterThan: f(constraint.greaterThan) }
  if ('lessThan' in constraint) return { lessThan: f(constraint.lessThan) }
  if ('greaterThanOrEqual' in constraint)
    return { greaterThanOrEqual: f(constraint.greaterThanOrEqual) }
  return { lessThanOrEqual: f(constraint.lessThanOrEqual) }
}

export function numberComparison(constraint: NumberConstraint): {
  operator: string
  formula1: number
  formula2?: number
} {
  if ('between' in constraint)
    return { operator: 'between', formula1: constraint.between[0], formula2: constraint.between[1] }
  if ('notBetween' in constraint)
    return {
      operator: 'notBetween',
      formula1: constraint.notBetween[0],
      formula2: constraint.notBetween[1],
    }
  if ('equal' in constraint) return { operator: 'equal', formula1: constraint.equal }
  if ('notEqual' in constraint) return { operator: 'notEqual', formula1: constraint.notEqual }
  if ('greaterThan' in constraint)
    return { operator: 'greaterThan', formula1: constraint.greaterThan }
  if ('lessThan' in constraint) return { operator: 'lessThan', formula1: constraint.lessThan }
  if ('greaterThanOrEqual' in constraint)
    return { operator: 'greaterThanOrEqual', formula1: constraint.greaterThanOrEqual }
  return { operator: 'lessThanOrEqual', formula1: constraint.lessThanOrEqual }
}

/** The inverse of `numberComparison`: an operator and its bounds back to a
 * constraint, or undefined for an operator this does not model. */
export function comparisonToConstraint(
  operator: string | undefined,
  formula1: string,
  formula2: string | undefined,
): NumberConstraint | undefined {
  const first = Number(formula1)
  if (!Number.isFinite(first)) return undefined
  switch (operator) {
    case 'equal':
      return { equal: first }
    case 'notEqual':
      return { notEqual: first }
    case 'greaterThan':
      return { greaterThan: first }
    case 'lessThan':
      return { lessThan: first }
    case 'greaterThanOrEqual':
      return { greaterThanOrEqual: first }
    case 'lessThanOrEqual':
      return { lessThanOrEqual: first }
    case 'between':
    case 'notBetween': {
      const second = Number(formula2)
      if (!Number.isFinite(second)) return undefined
      return operator === 'between' ? { between: [first, second] } : { notBetween: [first, second] }
    }
    default:
      return undefined
  }
}
