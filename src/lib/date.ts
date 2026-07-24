import { type XlsxErrorContext, XlsxError } from './errors.js'

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
export const LAST_SERIAL = 2_958_465

/**
 * A spreadsheet date is a calendar date rather than an instant, so a serial
 * means the same wall clock day wherever it is read, and `new Date(2024, 0, 1)`
 * writes the serial for that day.
 *
 * Both directions convert through the date's components rather than its
 * timestamp. Arithmetic on timestamps drifts by an hour whenever the date and
 * the epoch sit on opposite sides of a daylight saving change.
 */
export function serialToDate(serial: number, date1904: boolean, at: XlsxErrorContext = {}): Date {
  if (!Number.isFinite(serial) || serial < 0 || serial > LAST_SERIAL) {
    throw new XlsxError('invalid-content', `Serial ${serial} is not a date`, at)
  }

  const corrected = !date1904 && serial > PHANTOM_LEAP_DAY ? serial - 1 : serial
  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  // A serial is a count of days, so the product is fractional and the Date
  // constructor truncates it. Rounding recovers the exact millisecond a serial
  // written from a Date came from, and costs nothing on one that was not: the
  // format cannot express a serial finer than a Date can hold anyway.
  const asUtc = new Date(Math.round(epoch + corrected * MILLISECONDS_PER_DAY))

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

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/

/**
 * A `t="d"` cell holds an ISO-8601 literal. A spreadsheet date is a wall-clock
 * calendar date, so its components are read as written and any timezone
 * designator is ignored: `new Date(iso)` anchors a date-only or `Z`-suffixed
 * value to UTC, which shifts the day for a reader west of it. Returns undefined
 * when the text is not a real calendar date.
 */
export function parseIsoDate(text: string): Date | undefined {
  const match = ISO_DATE.exec(text)
  if (match === null) return undefined
  const [, year, month, day, hours, minutes, seconds, fraction] = match
  const milliseconds = fraction === undefined ? 0 : Math.round(Number(`0.${fraction}`) * 1000)
  const date = new Date(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hours ?? 0),
    Number(minutes ?? 0),
    Number(seconds ?? 0),
    milliseconds,
  )
  // A rolled-over field is not the date it was written as, and each overflow
  // spills into a larger unit: a day past the month's end bumps the month, an
  // hour past 23 the day, a second past 59 the minute, a fraction that rounds to
  // a full second likewise. Comparing every supplied component back rejects them
  // all, a bad day included. A year below 100, which the Date constructor maps
  // into the 1900s, fails the year check and is left as text.
  if (
    date.getFullYear() !== Number(year) ||
    date.getMonth() !== Number(month) - 1 ||
    date.getHours() !== Number(hours ?? 0) ||
    date.getMinutes() !== Number(minutes ?? 0) ||
    date.getSeconds() !== Number(seconds ?? 0)
  ) {
    return undefined
  }
  return date
}

export function dateToSerial(date: Date, date1904: boolean, at: XlsxErrorContext = {}): number {
  const where = at.reference === undefined ? '' : ` to cell ${at.reference}`

  if (Number.isNaN(date.getTime())) {
    throw new XlsxError('unwritable-value', `Cannot write an invalid date${where}`, at)
  }

  // Some zones move the clock forward at midnight, so the first moment of a
  // day is 01:00 there. A date sitting on it means that whole day, not an hour
  // past it, and has to give a whole serial.
  const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate())
  const atStartOfDay = date.getTime() === startOfDay.getTime()

  const asUtc = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    atStartOfDay ? 0 : date.getHours(),
    atStartOfDay ? 0 : date.getMinutes(),
    atStartOfDay ? 0 : date.getSeconds(),
    atStartOfDay ? 0 : date.getMilliseconds(),
  )

  const epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  const days = (asUtc - epoch) / MILLISECONDS_PER_DAY
  if (days < 0) {
    throw new XlsxError(
      'unwritable-value',
      `Cannot write ${date.toDateString()}${where}: it is before ` +
        `${date1904 ? 1904 : 1900}, which this workbook cannot hold`,
      at,
    )
  }

  const serial = date1904 ? days : days >= PHANTOM_LEAP_DAY ? days + 1 : days
  if (serial > LAST_SERIAL) {
    throw new XlsxError(
      'unwritable-value',
      `Cannot write ${date.toDateString()}${where}: it is after 9999, ` +
        'which this workbook cannot hold',
      at,
    )
  }
  return serial
}
