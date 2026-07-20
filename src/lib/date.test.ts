import assert from 'node:assert/strict'
import { test } from 'node:test'
import { serialToDate } from './date.js'

const iso = (serial: number, date1904 = false) => serialToDate(serial, date1904).toISOString()

test('reads the first day of the 1900 system', () => {
  assert.equal(iso(1), '1900-01-01T00:00:00.000Z')
})

test('reads a day before the phantom leap day', () => {
  assert.equal(iso(59), '1900-02-28T00:00:00.000Z')
})

test('reads a day after the phantom leap day', () => {
  assert.equal(iso(61), '1900-03-01T00:00:00.000Z')
})

test('maps the phantom leap day onto the day that follows it', () => {
  // Serial 60 is 1900-02-29 in Excel, a date that never existed.
  assert.equal(iso(60), '1900-03-01T00:00:00.000Z')
})

test('reads a modern date', () => {
  assert.equal(iso(45292), '2024-01-01T00:00:00.000Z')
})

test('reads the time of day from the fraction', () => {
  assert.equal(iso(1.5), '1900-01-01T12:00:00.000Z')
  assert.equal(iso(45292.25), '2024-01-01T06:00:00.000Z')
})

test('reads the first day of the 1904 system', () => {
  assert.equal(iso(0, true), '1904-01-01T00:00:00.000Z')
})

test('reads a later day of the 1904 system', () => {
  assert.equal(iso(1, true), '1904-01-02T00:00:00.000Z')
})

test('has no phantom leap day in the 1904 system', () => {
  assert.equal(iso(59, true), '1904-02-29T00:00:00.000Z')
})

test('reads the same day differently under each system', () => {
  assert.notEqual(iso(40000), iso(40000, true))
})

test('rejects a negative serial', () => {
  assert.throws(() => serialToDate(-1, false), /-1 is not a date/)
})

test('rejects a serial that is not a number', () => {
  assert.throws(() => serialToDate(Number.NaN, false), /NaN is not a date/)
})
