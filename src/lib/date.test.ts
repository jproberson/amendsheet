import assert from 'node:assert/strict'
import { test } from 'node:test'
import { dateToSerial, parseIsoDate, serialToDate } from './date.js'

/** Dates are wall clock, so they are asserted with local components. */
const stamp = (serial: number, date1904 = false) => {
  const date = serialToDate(serial, date1904)
  const pad = (value: number, width = 2) => String(value).padStart(width, '0')
  return (
    `${pad(date.getFullYear(), 4)}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  )
}
const iso = stamp

test('reads the first day of the 1900 system', () => {
  assert.equal(iso(1), '1900-01-01T00:00:00')
})

test('reads a day before the phantom leap day', () => {
  assert.equal(iso(59), '1900-02-28T00:00:00')
})

test('reads a day after the phantom leap day', () => {
  assert.equal(iso(61), '1900-03-01T00:00:00')
})

test('maps the phantom leap day onto the day that follows it', () => {
  // Serial 60 is 1900-02-29 in Excel, a date that never existed.
  assert.equal(iso(60), '1900-03-01T00:00:00')
})

test('reads a modern date', () => {
  assert.equal(iso(45292), '2024-01-01T00:00:00')
})

test('reads the time of day from the fraction', () => {
  assert.equal(iso(1.5), '1900-01-01T12:00:00')
  assert.equal(iso(45292.25), '2024-01-01T06:00:00')
})

test('reads the first day of the 1904 system', () => {
  assert.equal(iso(0, true), '1904-01-01T00:00:00')
})

test('reads a later day of the 1904 system', () => {
  assert.equal(iso(1, true), '1904-01-02T00:00:00')
})

test('has no phantom leap day in the 1904 system', () => {
  assert.equal(iso(59, true), '1904-02-29T00:00:00')
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

test('writes a date back as the serial it came from', () => {
  for (const serial of [1, 59, 61, 45292, 45292.25]) {
    assert.equal(dateToSerial(serialToDate(serial, false), false), serial)
  }
})

test('writes a date back under the 1904 epoch', () => {
  for (const serial of [0, 1, 59, 40000]) {
    assert.equal(dateToSerial(serialToDate(serial, true), true), serial)
  }
})

test('writes the phantom leap day as the day it maps onto', () => {
  assert.equal(dateToSerial(serialToDate(60, false), false), 61)
})

test('rejects writing an invalid date', () => {
  assert.throws(() => dateToSerial(new Date('nonsense'), false), /invalid date/)
})

test('a date built the ordinary way writes the serial for that day', () => {
  assert.equal(dateToSerial(new Date(2024, 0, 1), false), 45292)
})

test('a serial reads back as that calendar day locally', () => {
  const read = serialToDate(45292, false)

  assert.equal(read.getFullYear(), 2024)
  assert.equal(read.getMonth(), 0)
  assert.equal(read.getDate(), 1)
})

test('a time of day survives the round trip', () => {
  const noon = new Date(2024, 0, 1, 12, 30)

  assert.equal(serialToDate(dateToSerial(noon, false), false).getTime(), noon.getTime())
})

test('a summer date is a whole serial despite daylight saving', () => {
  for (const month of [0, 3, 6, 9]) {
    const serial = dateToSerial(new Date(2024, month, 15), false)
    assert.ok(Number.isInteger(serial), `month ${month} gave ${serial}`)
  }
})

test('every day of a year round trips to the same calendar day', () => {
  const day = new Date(2024, 0, 1)
  for (let index = 0; index < 366; index++) {
    const read = serialToDate(dateToSerial(day, false), false)
    assert.equal(read.getFullYear(), day.getFullYear(), day.toDateString())
    assert.equal(read.getMonth(), day.getMonth(), day.toDateString())
    assert.equal(read.getDate(), day.getDate(), day.toDateString())
    day.setDate(day.getDate() + 1)
  }
})

test('refuses a date the workbook epoch cannot represent', () => {
  // The 1904 system has no way to write a day before 1904.
  assert.throws(() => dateToSerial(new Date(1900, 1, 28), true), /before/i)
})

test('writes the first day the 1904 system can hold', () => {
  assert.equal(dateToSerial(new Date(1904, 0, 1), true), 0)
})

test('refuses a serial beyond the last day a sheet can hold', () => {
  assert.throws(() => serialToDate(1e12, false), /not a date/)
  assert.throws(() => serialToDate(2958466, false), /not a date/)
})

test('reads the last day a sheet can hold', () => {
  assert.equal(serialToDate(2958465, false).getFullYear(), 9999)
})

test('a day whose midnight does not exist still writes a whole serial', () => {
  // Some zones move the clock forward at midnight, so new Date(y, m, d) is
  // 01:00 there. Each day is built fresh, the way a caller writes one.
  for (let month = 0; month < 12; month++) {
    for (let day = 1; day <= 28; day++) {
      const built = new Date(2024, month, day)
      const serial = dateToSerial(built, false)
      assert.ok(Number.isInteger(serial), `${built.toDateString()} gave ${serial}`)
      assert.equal(serialToDate(serial, false).getDate(), day, built.toDateString())
    }
  }
})

test('a date survives the serial round trip to the millisecond', () => {
  // corrected * MILLISECONDS_PER_DAY is fractional, and the Date constructor
  // truncates it, so a time of day came back a millisecond early.
  const fields = (date: Date) =>
    [
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds(),
      date.getMilliseconds(),
    ].join(' ')

  for (const date1904 of [false, true]) {
    for (let year = 1905; year < 2400; year += 7) {
      for (const [month, day, ms] of [
        [0, 5, 70],
        [5, 30, 999],
        [11, 31, 1],
      ]) {
        const source = new Date(year, month ?? 0, day ?? 1, 2, 35, 52, ms ?? 0)
        const back = serialToDate(dateToSerial(source, date1904), date1904)
        assert.equal(fields(back), fields(source), `${source.toISOString()} 1904=${date1904}`)
      }
    }
  }
})

test('an ISO date-only literal is the same wall-clock day in every timezone', () => {
  const date = parseIsoDate('2024-03-15')
  assert.ok(date !== undefined)
  assert.equal(date?.getDate(), 15)
  assert.equal(dateToSerial(date ?? new Date(Number.NaN), false), 45366)
})

test('an ISO datetime literal keeps its time of day', () => {
  const date = parseIsoDate('2024-03-15T12:00:00')

  assert.equal(dateToSerial(date ?? new Date(Number.NaN), false), 45366.5)
})

test('an ISO literal with a trailing Z is still read as its wall-clock day', () => {
  const date = parseIsoDate('2024-03-15T00:00:00Z')

  assert.equal(dateToSerial(date ?? new Date(Number.NaN), false), 45366)
})

test('an ISO literal carries a sub-second fraction', () => {
  const date = parseIsoDate('2024-03-15T00:00:00.250')

  assert.equal(date?.getMilliseconds(), 250)
})

test('a non-date literal is rejected', () => {
  assert.equal(parseIsoDate('not a date'), undefined)
  assert.equal(parseIsoDate('2024-02-30'), undefined)
  // Below 100 the Date constructor maps the year into the 1900s, so the year no
  // longer matches what was written and the literal is left as text.
  assert.equal(parseIsoDate('0050-01-01'), undefined)
})
