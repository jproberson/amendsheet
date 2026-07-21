import { XlsxError } from './errors.js'

const MILLISECONDS_PER_DAY = 86_400_000

/** Serial 1 is 1900-01-01, so day zero sits on the last day of 1899. */
const EPOCH_1900 = Date.UTC(1899, 11, 31)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

/**
 * Excel counts 1900 as a leap year for Lotus 1-2-3 compatibility, so serial 60
 * is 1900-02-29, a day that never existed, and every serial after it is one too
 * high.
 */
const PHANTOM_LEAP_DAY = 60

/** 9999-12-31, the last day a sheet can hold. */
const LAST_SERIAL = 2_958_465

/**
 * A spreadsheet date is a calendar date rather than an instant, so a serial
 * means the same wall clock day wherever it is read, and `new Date(2024, 0, 1)`
 * writes the serial for that day.
 *
 * Both directions convert through the date's components rather than its
 * timestamp. Arithmetic on timestamps drifts by an hour whenever the date and
 * the epoch sit on opposite sides of a daylight saving change.
 */
export function serialToDate(serial: number, date1904: boolean): Date {
  if (!Number.isFinite(serial) || serial < 0 || serial > LAST_SERIAL) {
    throw new XlsxError('unwritable-value', `Serial ${serial} is not a date`)
  }

  const corrected = !date1904 && serial > PHANTOM_LEAP_DAY ? serial - 1 : serial
  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  const asUtc = new Date(epoch + corrected * MILLISECONDS_PER_DAY)

  return new Date(
    asUtc.getUTCFullYear(),
    asUtc.getUTCMonth(),
    asUtc.getUTCDate(),
    asUtc.getUTCHours(),
    asUtc.getUTCMinutes(),
    asUtc.getUTCSeconds(),
    asUtc.getUTCMilliseconds(),
  )
}

export function dateToSerial(date: Date, date1904: boolean): number {
  if (Number.isNaN(date.getTime())) {
    throw new XlsxError('unwritable-value', 'Invalid date cannot be written')
  }

  const asUtc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds(),
  )

  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  const days = (asUtc - epoch) / MILLISECONDS_PER_DAY
  if (days < 0) {
    throw new XlsxError(
      'unwritable-value',
      `${date.toDateString()} is before ${date1904 ? 1904 : 1900}, which this workbook cannot hold`,
    )
  }

  if (date1904) return days
  return days >= PHANTOM_LEAP_DAY ? days + 1 : days
}
