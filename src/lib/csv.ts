/**
 * Delimited-text parsing and formatting, following RFC 4180: fields are separated
 * by `delimiter`, a field holding the delimiter, a quote or a line break is wrapped
 * in double quotes, and a quote inside such a field is doubled. Rows end with CRLF,
 * LF or a lone CR on read; formatting writes CRLF. Pure string in, string out, so
 * it is safe in a browser and holds no spreadsheet knowledge of its own.
 */
export function parseCsv(text: string, delimiter = ','): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let inQuotes = false
  let started = false
  let index = 0
  while (index < text.length) {
    const character = text[index]
    if (inQuotes) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 2
          continue
        }
        inQuotes = false
        index += 1
        continue
      }
      field += character
      index += 1
      continue
    }
    if (character === '"') {
      inQuotes = true
      started = true
      index += 1
      continue
    }
    if (character === delimiter) {
      row.push(field)
      field = ''
      started = true
      index += 1
      continue
    }
    if (character === '\n' || character === '\r') {
      row.push(field)
      rows.push(row)
      field = ''
      row = []
      started = false
      index += character === '\r' && text[index + 1] === '\n' ? 2 : 1
      continue
    }
    field += character
    started = true
    index += 1
  }
  // A row still building at the end had no trailing newline, so it is the last row;
  // a file that ended on a newline leaves nothing started, so no empty row is added.
  if (started || field !== '' || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

export function formatCsv(rows: ReadonlyArray<ReadonlyArray<string>>, delimiter = ','): string {
  const needsQuoting = (field: string): boolean =>
    field.includes(delimiter) || field.includes('"') || field.includes('\n') || field.includes('\r')
  const quoted = (field: string): string =>
    needsQuoting(field) ? `"${field.replaceAll('"', '""')}"` : field
  return rows.map((row) => row.map(quoted).join(delimiter)).join('\r\n')
}
