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

import { escapeXml, extractElements } from '../shared/xml.js'

/**
 * Read an attribute from an opening tag, matching the whole attribute name.
 *
 * A bare `name="…"` search is not enough: `width` also appears inside
 * `customWidth`, and `s` inside `spans`. R7's own writer emits
 * `<col customWidth="1" min="1" max="1" width="14"/>`, so an unanchored search
 * reports the width of a column as 1 — the value of `customWidth`. The leading
 * boundary (start of tag or whitespace) makes the match exact.
 *
 * @param {string} tagXml
 * @param {string} name
 * @returns {string|null}
 */
export function attr(tagXml, name) {
  if (!tagXml) return null
  const escaped = String(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = String(tagXml).match(new RegExp(`(?:^|[\\s<])${escaped}\\s*=\\s*["']([^"']*)["']`))
  return match ? match[1] : null
}

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
  if (!sheetData) return { open: '<sheetData>', rows: new Map(), present: false }

  // A self-closing `<sheetData/>` cannot host rows, so the reopened form is
  // used as the container tag from here on.
  const openTag = sheetData.outerXml.match(/^<sheetData[^>]*>/)?.[0] || '<sheetData>'
  const rows = new Map()
  for (const row of extractElements(sheetData.innerXml, 'row')) {
    const r = Number(attr(row.outerXml, 'r'))
    if (Number.isFinite(r)) rows.set(r, row.outerXml)
  }
  return {
    open: openTag.endsWith('/>') ? '<sheetData>' : openTag,
    rows,
    present: true
  }
}

/** Re-emit `<sheetData>` with rows in ascending order. */
function renderRows(parsed) {
  const numbers = [...parsed.rows.keys()].sort((a, b) => a - b)
  const body = numbers.map((n) => parsed.rows.get(n)).join('')
  return `${parsed.open}${body}</sheetData>`
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
    const ref = attr(c.outerXml, 'r')
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

/**
 * Row-level attributes worth preserving when a row is rebuilt.
 *
 * Everything except `r` is kept: a cell edit must not discard the row's height,
 * its hidden flag or any attribute this module does not model. Only
 * {@link setRowHeight} asks for the height to be dropped, because it is about
 * to write a new one.
 */
function rowAttributes(rowXml, { dropHeight = false } = {}) {
  if (!rowXml) return ''
  const open = rowXml.match(/^<row[^>]*>/)
  if (!open) return ''
  let attrs = open[0]
    .replace(/^<row/, '')
    .replace(/>$/, '')
    .replace(/\s*r="[^"]*"/, '')
  if (dropHeight) {
    attrs = attrs
      .replace(/\s*ht="[^"]*"/, '')
      .replace(/\s*customHeight="[^"]*"/, '')
  }
  return attrs
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
  // Content cannot live inside a self-closing tag, so a cell that had no
  // payload until now is reopened before the payload is written.
  const cellOpenTag = updated.inner && updated.openTag.endsWith('/>')
    ? `${updated.openTag.slice(0, -2)}>`
    : updated.openTag
  const cellXml = updated.inner
    ? `${cellOpenTag}${updated.inner}</c>`
    : cellOpenTag.endsWith('/>')
      ? cellOpenTag
      : `${cellOpenTag}</c>`

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

/**
 * Replace a cell's payload while keeping its style and every attribute this
 * module does not model.
 *
 * Writing a value must not silently unformat the cell, and it must not reorder
 * the row: both are handled here because the cell element is looked up and
 * re-emitted in place.
 *
 * @param {string} sheetXml
 * @param {string} ref
 * @param {{content?: string, type?: string|null, xfIndex?: number|null}} spec
 *   `type` is the OOXML `t` attribute (`inlineStr`, `str`, …) or null to drop
 *   it; `xfIndex` is only applied when provided (null or 0 clears the style).
 * @returns {string}
 */
export function setCellContent(sheetXml, ref, spec = {}) {
  const { content = '', type = null, xfIndex } = spec
  return withCell(sheetXml, ref, (openTag) => {
    let tag = removeAttr(openTag, 't')
    if (type) tag = setAttr(tag, 't', type)
    if (xfIndex !== undefined) {
      tag = xfIndex === null || xfIndex === 0
        ? removeAttr(tag, 's')
        : setAttr(tag, 's', xfIndex)
    }
    return { openTag: tag, inner: content }
  })
}

/** The cell format index of a cell, or null when it has none. */
export function getCellStyle(sheetXml, ref) {
  const { row, col } = parseRef(ref)
  const parsed = parseRows(sheetXml)
  const rowXml = parsed.rows.get(row + 1)
  if (!rowXml) return null
  const cell = parseRowCells(rowXml).get(col)
  if (!cell) return null
  const s = attr(cell, 's')
  return s === null || s === undefined ? null : Number(s)
}

/** Every merged range declared in the worksheet. */
export function getMergedCells(sheetXml) {
  const container = extractElements(sheetXml, 'mergeCells')[0]
  if (!container) return []
  return extractElements(container.outerXml, 'mergeCell')
    .map((m) => attr(m.outerXml, 'ref'))
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

/**
 * The used range of the worksheet, derived from the cells actually present.
 * @returns {string|null} an A1 or A1:B2 reference, or null for an empty sheet.
 */
export function getUsedRange(sheetXml) {
  const parsed = parseRows(sheetXml)
  let minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1
  for (const [rowNum, rowXml] of parsed.rows) {
    for (const col of parseRowCells(rowXml).keys()) {
      if (rowNum < minRow) minRow = rowNum
      if (rowNum > maxRow) maxRow = rowNum
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    }
  }
  if (maxRow === -1) return null
  const first = buildRef(minRow - 1, minCol)
  const last = buildRef(maxRow - 1, maxCol)
  return first === last ? first : `${first}:${last}`
}

/**
 * Keep `<dimension>` honest after structural edits.
 *
 * The hint is recomputed from the cells that are really there rather than
 * blanked, and it is left alone when it is already correct — an edit must not
 * churn a part more than the caller asked for.
 */
function replaceDimensions(sheetXml) {
  const dimension = extractElements(sheetXml, 'dimension')[0]
  if (!dimension) return sheetXml
  const bounds = getUsedRange(sheetXml) || 'A1'
  if (attr(dimension.outerXml, 'ref') === bounds) return sheetXml
  const updated = dimension.outerXml.replace(/ref=["'][^"']*["']/, `ref="${bounds}"`)
  return sheetXml.replace(dimension.outerXml, updated)
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
      // The original tag is kept so attributes this module does not model
      // (a column-level style, bestFit, …) survive a width change.
      raw: c.outerXml,
      min: Number(attr(c.outerXml, 'min')),
      max: Number(attr(c.outerXml, 'max')),
      width: Number(attr(c.outerXml, 'width')),
      customWidth: attr(c.outerXml, 'customWidth')
    }))
    : []

  // Split any existing span that overlaps the requested one, so unrelated
  // columns keep the width the author gave them.
  const rebuilt = []
  for (const col of cols) {
    if (!Number.isFinite(col.min) || !Number.isFinite(col.max)) {
      // A span we cannot interpret is carried through rather than mangled.
      rebuilt.push(col)
      continue
    }
    if (col.max < min || col.min > max) {
      rebuilt.push(col)
      continue
    }
    if (col.min < min) rebuilt.push({ ...col, max: min - 1 })
    if (col.max > max) rebuilt.push({ ...col, min: max + 1 })
  }
  rebuilt.push({ min, max, width, customWidth: '1', raw: null })
  rebuilt.sort((a, b) => a.min - b.min)

  const xml = '<cols>'
    + rebuilt.map((c) => {
      const tag = `<col min="${c.min}" max="${c.max}" width="${c.width}" customWidth="${c.customWidth ?? '1'}"/>`
      if (!c.raw) return tag
      return c.raw
        .replace(/\bmin="[^"]*"/, `min="${c.min}"`)
        .replace(/\bmax="[^"]*"/, `max="${c.max}"`)
    }).join('')
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
    min: Number(attr(c.outerXml, 'min')),
    max: Number(attr(c.outerXml, 'max')),
    width: Number(attr(c.outerXml, 'width'))
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
  const attrs = rowAttributes(existing, { dropHeight: true })
  const cells = existing ? parseRowCells(existing) : new Map()
  parsed.rows.set(rowNumber, `<row r="${rowNumber}"${attrs} ht="${height}" customHeight="1">${[...cells.keys()].sort((a, b) => a - b).map((c) => cells.get(c)).join('')}</row>`)
  return replaceSheetData(sheetXml, renderRows(parsed))
}

/** Row heights declared in the worksheet, as `{row: height}`. */
export function getRowHeights(sheetXml) {
  const parsed = parseRows(sheetXml)
  const out = {}
  for (const [num, xml] of parsed.rows) {
    const ht = attr(xml, 'ht')
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
