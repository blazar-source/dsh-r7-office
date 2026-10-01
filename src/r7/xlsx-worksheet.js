/**
 * Worksheet-level structural edits for XLSX.
 *
 * Everything here is a pure function from sheet XML to sheet XML. The rules
 * that make the operations safe on a file this project did not author:
 *
 *   - existing `<row>` and `<c>` elements are kept verbatim; only the
 *     attributes being changed are touched;
 *   - rows are re-emitted in ascending `r` order and cells in ascending column
 *     order, because Excel and R7 reject a worksheet that is out of order —
 *     which is also why inserting a cell is an ordered insert, not an append;
 *   - a style is applied by setting `s` on the cell, never by rewriting the
 *     cell's value, type or formula.
 */

import { escapeXml, getAttribute, extractElements } from '../shared/xml.js'

/** Convert a column letter to a 0-based index (`A` -> 0, `AA` -> 26). */
export function colToIndex(letters) {
  let index = 0
  for (const ch of letters.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64)
  }
  return index - 1
}

/** Convert a 0-based column index to letters (`0` -> `A`). */
export function indexToCol(index) {
  let out = ''
  let n = index + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/**
 * Parse an A1 reference.
 * @param {string} ref
 * @returns {{col: number, row: number, colLetters: string}} row is 0-based.
 */
export function parseRef(ref) {
  const match = String(ref).trim().match(/^\$?([A-Za-z]+)\$?(\d+)$/)
  if (!match) throw new Error(`Invalid cell reference: "${ref}"`)
  return {
    colLetters: match[1].toUpperCase(),
    col: colToIndex(match[1]),
    row: Number(match[2]) - 1
  }
}

/** Build an A1 reference from a 0-based row and column. */
export function buildRef(row, col) {
  return `${indexToCol(col)}${row + 1}`
}

/**
 * Parse a range like `A1:D10`, `B2` or `A1:D10` with `$` anchors.
 * @param {string} range
 * @returns {{start: object, end: object}}
 */
export function parseRange(range) {
  const parts = String(range).trim().split(':')
  const start = parseRef(parts[0])
  const end = parts[1] ? parseRef(parts[1]) : { ...start }
  if (end.row < start.row || end.col < start.col) {
    throw new Error(`Invalid range "${range}": the end must not precede the start`)
  }
  return { start, end }
}

/** Every cell reference inside a range, row-major. */
export function expandRange(range) {
  const { start, end } = parseRange(range)
  const refs = []
  for (let r = start.row; r <= end.row; r++) {
    for (let c = start.col; c <= end.col; c++) refs.push(buildRef(r, c))
  }
  return refs
}

/** Normalize a range to `A1:D10` form. */
export function normalizeRange(range) {
  const { start, end } = parseRange(range)
  const a = buildRef(start.row, start.col)
  const b = buildRef(end.row, end.col)
  return a === b ? a : `${a}:${b}`
}

/** Split `<sheetData>` into ordered row XML strings, keyed by row number. */
function parseRows(sheetXml) {
  const sheetData = extractElements(sheetXml, 'sheetData')[0]
  if (!sheetData) return { open: '<sheetData>', close: '</sheetData>', rows: new Map(), present: false }

  const inner = sheetData.innerXml
  const rows = new Map()
  for (const row of extractElements(inner, 'row')) {
    const r = Number(getAttribute(row.outerXml, 'r'))
    if (Number.isFinite(r)) rows.set(r, row.outerXml)
  }
  return { open: '<sheetData>', close: '</sheetData>', rows, present: true }
}

/** Re-emit `<sheetData>` with rows in ascending order. */
function renderRows(parsed) {
  const numbers = [...parsed.rows.keys()].sort((a, b) => a - b)
  const body = numbers.map((n) => parsed.rows.get(n)).join('')
  return `<sheetData>${body}</sheetData>`
}

/** Replace the `<sheetData>` element with a rebuilt one. */
function replaceSheetData(sheetXml, rendered) {
  const existing = extractElements(sheetXml, 'sheetData')[0]
  if (existing) return sheetXml.replace(existing.outerXml, rendered)
  return sheetXml.replace('</worksheet>', `${rendered}</worksheet>`)
}

/** Cells of a row, in ascending column order. */
function parseRowCells(rowXml) {
  const cells = new Map()
  for (const c of extractElements(rowXml, 'c')) {
    const ref = getAttribute(c.outerXml, 'r')
    if (!ref) continue
    cells.set(colToIndex(ref.replace(/[^A-Za-z]/g, '')), c.outerXml)
  }
  return cells
}

/** Re-emit a row with its cells in ascending column order. */
function renderRow(rowNum, rowAttrs, cells) {
  const cols = [...cells.keys()].sort((a, b) => a - b)
  const body = cols.map((c) => cells.get(c)).join('')
  return `<row r="${rowNum}"${rowAttrs}>${body}</row>`
}

/** Row-level attributes worth preserving when a row is rebuilt. */
function rowAttributes(rowXml) {
  if (!rowXml) return ''
  const open = rowXml.match(/^<row[^>]*>/)
  if (!open) return ''
  return open[0]
    .replace(/^<row/, '')
    .replace(/>$/, '')
    .replace(/\s*r="[^"]*"/, '')
    .replace(/\s*ht="[^"]*"/, '')
    .replace(/\s*customHeight="[^"]*"/, '')
    .replace(/\s*spans="[^"]*"/, '')
}

/**
 * Ensure a cell exists, then update the attributes the caller asks for.
 *
 * @param {string} sheetXml
 * @param {string} ref
 * @param {(openTag: string, inner: string) => {openTag: string, inner: string}} mutate
 * @returns {string} new sheet XML
 */
function withCell(sheetXml, ref, mutate) {
  const { row, col } = parseRef(ref)
  const rowNum = row + 1
  const parsed = parseRows(sheetXml)

  const existingRow = parsed.rows.get(rowNum)
  const attrs = rowAttributes(existingRow)
  const cells = existingRow ? parseRowCells(existingRow) : new Map()

  const existingCell = cells.get(col)
  let openTag = `<c r="${ref}"/>`
  let inner = ''
  if (existingCell) {
    const open = existingCell.match(/^<c[^>]*?(\/?)>/)
    if (open) {
      openTag = open[0]
      inner = existingCell.slice(open[0].length)
      if (inner.endsWith('</c>')) inner = inner.slice(0, -4)
    }
  }

  const updated = mutate(openTag, inner)
  const cellXml = updated.inner
    ? `${updated.openTag}${updated.inner}</c>`
    : updated.openTag.endsWith('/>')
      ? updated.openTag
      : `${updated.openTag}</c>`

  cells.set(col, cellXml)
  parsed.rows.set(rowNum, renderRow(rowNum, attrs, cells))
  return replaceSheetData(sheetXml, renderRows(parsed))
}

/** Set (or add) an attribute on a cell's opening tag. */
function setAttr(openTag, name, value) {
  const selfClosing = openTag.endsWith('/>')
  const bare = selfClosing ? openTag.slice(0, -2) : openTag.slice(0, -1)
  const existing = new RegExp(`\\s${name}="[^"]*"`)
  const next = existing.test(bare)
    ? bare.replace(existing, ` ${name}="${escapeXml(String(value))}"`)
    : `${bare} ${name}="${escapeXml(String(value))}"`
  return `${next}${selfClosing ? '/>' : '>'}`
}

/** Remove an attribute from a cell's opening tag. */
function removeAttr(openTag, name) {
  const selfClosing = openTag.endsWith('/>')
  const bare = selfClosing ? openTag.slice(0, -2) : openTag.slice(0, -1)
  const next = bare.replace(new RegExp(`\\s${name}="[^"]*"`), '')
  return `${next}${selfClosing ? '/>' : '>'}`
}

/**
 * Point a cell at a cell format (`<c s="…">`).
 * @param {string} sheetXml
 * @param {string} ref
 * @param {number|null} xfIndex - null removes the style (back to the default).
 * @returns {string}
 */
export function setCellStyle(sheetXml, ref, xfIndex) {
  return withCell(sheetXml, ref, (openTag, inner) => ({
    openTag: xfIndex === null || xfIndex === 0
      ? removeAttr(openTag, 's')
      : setAttr(openTag, 's', xfIndex),
    inner
  }))
}

/**
 * Ensure a cell element exists without changing its style.
 * @param {string} sheetXml
 * @param {string} ref
 * @returns {string}
 */
export function ensureCell(sheetXml, ref) {
  return withCell(sheetXml, ref, (openTag, inner) => ({ openTag, inner }))
}

/** The cell format index of a cell, or null when it has none. */
export function getCellStyle(sheetXml, ref) {
  const { row, col } = parseRef(ref)
  const parsed = parseRows(sheetXml)
  const rowXml = parsed.rows.get(row + 1)
  if (!rowXml) return null
  const cell = parseRowCells(rowXml).get(col)
  if (!cell) return null
  const s = getAttribute(cell, 's')
  return s === null || s === undefined ? null : Number(s)
}

/** Every merged range declared in the worksheet. */
export function getMergedCells(sheetXml) {
  const container = extractElements(sheetXml, 'mergeCells')[0]
  if (!container) return []
  return extractElements(container.outerXml, 'mergeCell')
    .map((m) => getAttribute(m.outerXml, 'ref'))
    .filter(Boolean)
}

/**
 * Add and/or remove merged ranges.
 * @param {string} sheetXml
 * @param {{add?: string[], remove?: string[]}} changes
 * @returns {string}
 */
export function setMergedCells(sheetXml, changes = {}) {
  const add = (changes.add || []).map(normalizeRange)
  const remove = new Set((changes.remove || []).map(normalizeRange))
  const current = getMergedCells(sheetXml).filter((r) => !remove.has(normalizeRange(r)))
  const merged = [...current]
  for (const r of add) if (!merged.includes(r)) merged.push(r)

  const container = extractElements(sheetXml, 'mergeCells')[0]
  let updated = sheetXml
  if (container) updated = updated.replace(container.outerXml, '')

  if (merged.length === 0) return replaceDimensions(updated)

  const xml = `<mergeCells count="${merged.length}">`
    + merged.map((r) => `<mergeCell ref="${r}"/>`).join('')
    + '</mergeCells>'

  // mergeCells must follow sheetData in the worksheet schema.
  const sheetData = extractElements(updated, 'sheetData')[0]
  if (sheetData) {
    updated = updated.replace(sheetData.outerXml, `${sheetData.outerXml}${xml}`)
  } else {
    updated = updated.replace('</worksheet>', `${xml}</worksheet>`)
  }
  return replaceDimensions(updated)
}

/** Keep `<dimension>` roughly honest after structural edits. */
function replaceDimensions(sheetXml) {
  const dimension = extractElements(sheetXml, 'dimension')[0]
  if (!dimension) return sheetXml
  return sheetXml.replace(dimension.outerXml, '<dimension ref="A1"/>')
}

/**
 * Set a column width for a span of columns.
 * @param {string} sheetXml
 * @param {{min: number, max: number, width: number}} spec - 1-based columns.
 * @returns {string}
 */
export function setColumnWidth(sheetXml, spec) {
  const min = Number(spec.min)
  const max = Number(spec.max ?? spec.min)
  const width = Number(spec.width)
  if (!Number.isFinite(width) || width <= 0) {
    throw new Error(`Invalid column width: ${spec.width}`)
  }

  const container = extractElements(sheetXml, 'cols')[0]
  const cols = container
    ? extractElements(container.outerXml, 'col').map((c) => ({
      min: Number(getAttribute(c.outerXml, 'min')),
      max: Number(getAttribute(c.outerXml, 'max')),
      width: Number(getAttribute(c.outerXml, 'width')),
      customWidth: getAttribute(c.outerXml, 'customWidth')
    }))
    : []

  // Split any existing span that overlaps the requested one, so unrelated
  // columns keep the width the author gave them.
  const rebuilt = []
  for (const col of cols) {
    if (col.max < min || col.min > max) {
      rebuilt.push(col)
      continue
    }
    if (col.min < min) rebuilt.push({ ...col, max: min - 1 })
    if (col.max > max) rebuilt.push({ ...col, min: max + 1 })
  }
  rebuilt.push({ min, max, width, customWidth: '1' })
  rebuilt.sort((a, b) => a.min - b.min)

  const xml = '<cols>'
    + rebuilt.map((c) => `<col min="${c.min}" max="${c.max}" width="${c.width}" customWidth="${c.customWidth ?? '1'}"/>`).join('')
    + '</cols>'

  let updated = container ? sheetXml.replace(container.outerXml, xml) : sheetXml
  if (!container) {
    // cols must precede sheetData.
    const sheetData = extractElements(updated, 'sheetData')[0]
    if (sheetData) {
      updated = updated.replace(sheetData.outerXml, `${xml}${sheetData.outerXml}`)
    } else {
      updated = updated.replace('</worksheet>', `${xml}</worksheet>`)
    }
  }
  return replaceDimensions(updated)
}

/** Column widths declared in the worksheet, as `{min,max,width}` records. */
export function getColumnWidths(sheetXml) {
  const container = extractElements(sheetXml, 'cols')[0]
  if (!container) return []
  return extractElements(container.outerXml, 'col').map((c) => ({
    min: Number(getAttribute(c.outerXml, 'min')),
    max: Number(getAttribute(c.outerXml, 'max')),
    width: Number(getAttribute(c.outerXml, 'width'))
  }))
}

/**
 * Set a row height.
 * @param {string} sheetXml
 * @param {number} rowNumber - 1-based
 * @param {number} height - in points
 * @returns {string}
 */
export function setRowHeight(sheetXml, rowNumber, height) {
  if (!Number.isFinite(height) || height <= 0) {
    throw new Error(`Invalid row height: ${height}`)
  }
  const parsed = parseRows(sheetXml)
  const existing = parsed.rows.get(rowNumber)
  const attrs = rowAttributes(existing)
  const cells = existing ? parseRowCells(existing) : new Map()
  parsed.rows.set(rowNumber, `<row r="${rowNumber}"${attrs} ht="${height}" customHeight="1">${[...cells.keys()].sort((a, b) => a - b).map((c) => cells.get(c)).join('')}</row>`)
  return replaceSheetData(sheetXml, renderRows(parsed))
}

/** Row heights declared in the worksheet, as `{row: height}`. */
export function getRowHeights(sheetXml) {
  const parsed = parseRows(sheetXml)
  const out = {}
  for (const [num, xml] of parsed.rows) {
    const ht = getAttribute(xml, 'ht')
    if (ht !== null && ht !== undefined) out[num] = Number(ht)
  }
  return out
}

/**
 * Estimate a column width that fits the longest value in a column.
 *
 * This is a heuristic: the exact rendered width depends on the font, which a
 * spreadsheet engine resolves at layout time. The estimate is deliberately
 * generous so a value is not clipped.
 *
 * @param {Array<{text: string}>} values
 * @param {number} [minWidth=8]
 * @returns {number}
 */
export function estimateColumnWidth(values, minWidth = 8) {
  let longest = 0
  for (const v of values) {
    const text = String(v?.text ?? '')
    // Count a wide glyph as roughly two narrow ones.
    let width = 0
    for (const ch of text) width += /[\u3000-\u9fff\uff00-\uffef]/.test(ch) ? 2 : 1
    if (width > longest) longest = width
  }
  return Math.min(Math.max(longest + 2, minWidth), 255)
}
