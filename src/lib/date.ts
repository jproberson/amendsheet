import { XlsxError } from './errors.js'

const MILLISECONDS_PER_DAY = 86_400_000

/** Serial 1 is 1900-01-01, so day zero sits on the last day of 1899. */
const EPOCH_1900 = Date.UTC(1899, 11, 31)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

/**
 * Excel counts 1900 as a leap year to stay compatible with Lotus 1-2-3, so
 * serial 60 is 1900-02-29 — a day that never existed. Every serial after it is
 * one greater than the true day count, which this corrects.
 */
const PHANTOM_LEAP_DAY = 60

/**
 * Turns a stored serial into a date. Which epoch applies is a property of the
 * workbook, not the cell, so it has to be passed in.
 */
export function serialToDate(serial: number, date1904: boolean): Date {
  if (!Number.isFinite(serial) || serial < 0) {
    throw new XlsxError(`Serial ${serial} is not a date`)
  }

  if (date1904) {
    return new Date(EPOCH_1904 + serial * MILLISECONDS_PER_DAY)
  }

  const corrected = serial > PHANTOM_LEAP_DAY ? serial - 1 : serial
  return new Date(EPOCH_1900 + corrected * MILLISECONDS_PER_DAY)
}
