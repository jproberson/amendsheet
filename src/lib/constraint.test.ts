import assert from 'node:assert/strict'
import { test } from 'node:test'
import { comparisonToConstraint, mapConstraint, numberComparison } from './constraint.js'
import type { NumberConstraint } from './public-types.js'

test('numberComparison names a single-bound operator', () => {
  assert.deepEqual(numberComparison({ greaterThan: 5 }), { operator: 'greaterThan', formula1: 5 })
})

test('numberComparison carries both bounds of a between', () => {
  assert.deepEqual(numberComparison({ between: [1, 9] }), {
    operator: 'between',
    formula1: 1,
    formula2: 9,
  })
})

test('comparisonToConstraint inverts numberComparison for every operator', () => {
  const cases: NumberConstraint[] = [
    { equal: 3 },
    { notEqual: 3 },
    { greaterThan: 3 },
    { lessThan: 3 },
    { greaterThanOrEqual: 3 },
    { lessThanOrEqual: 3 },
    { between: [2, 8] },
    { notBetween: [2, 8] },
  ]
  for (const constraint of cases) {
    const c = numberComparison(constraint)
    assert.deepEqual(
      comparisonToConstraint(
        c.operator,
        String(c.formula1),
        c.formula2 === undefined ? undefined : String(c.formula2),
      ),
      constraint,
    )
  }
})

test('comparisonToConstraint rejects an unknown operator or a non-finite bound', () => {
  assert.equal(comparisonToConstraint('atMost', '5', undefined), undefined)
  assert.equal(comparisonToConstraint('equal', 'not-a-number', undefined), undefined)
  assert.equal(comparisonToConstraint('between', '1', 'nope'), undefined)
})

test('mapConstraint transforms both bounds and preserves the operator', () => {
  assert.deepEqual(
    mapConstraint({ between: [1, 2] }, (n) => n * 10),
    { between: [10, 20] },
  )
  assert.deepEqual(
    mapConstraint({ lessThanOrEqual: 4 }, (n) => n + 1),
    { lessThanOrEqual: 5 },
  )
})
