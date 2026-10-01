import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'
import {
  Stylesheet,
  dateToSerial,
  normalizeArgb,
  resolveNumberFormat
} from './xlsx-styles.js'
import {
  colToIndex,
  indexToCol,
  expandRange,
  normalizeRange,
  parseRange,
  parseRef,
  attr,
  buildRef,
  getCellStyle,
  setCellStyle,
  setCellContent,
  getMergedCells,
  setMergedCells,
  getColumnWidths,
  setColumnWidth,
  getRowHeights,
  setRowHeight,
  estimateColumnWidth
} from './xlsx-worksheet.js'

/**
 * Column/row reference helpers.
 *
 * These delegate to the worksheet module so a single implementation decides how
 * a reference is parsed; two copies would eventually disagree about `$`
 * anchors or multi-letter columns.
 */
export function colLetterToIndex(colStr) {
  return colToIndex(colStr)
}

export function indexToColLetter(index) {
  return indexToCol(index)
}

export function parseCellRef(cellRef) {
  const parsed = parseRef(cellRef)
  return {
    colStr: parsed.colLetters,
    col: parsed.col,
    row: parsed.row,
    rowNum: parsed.row + 1
  }
}

/** Whether an ISO 8601 value carries a time part, i.e. needs a datetime format. */
function hasTimePart(value) {
  const text = value instanceof Date ? value.toISOString() : String(value)
  const time = text.includes('T') ? text.slice(text.indexOf('T') + 1) : text
  return /\d{1,2}:\d{2}/.test(time)
}

/**
 * Expand the friendly border request into per-edge entries.
 *
 * `{ all: 'thin', top: null }` means "every edge thin except the top" — the
 * worksheet module only understands edges, so `all` is resolved here.
 *
 * @param {object} border
 * @returns {object|null} per-edge spec, or null when nothing is requested.
 */
function expandBorderSpec(border) {
  const out = {}
  for (const edge of ['left', 'right', 'top', 'bottom']) {
    const value = border[edge] !== undefined ? border[edge] : border.all
    if (value === undefined || value === null || value === '') continue
    out[edge] = value
  }
  return Object.keys(out).length > 0 ? out : null
}

/** Whether two A1 ranges share at least one cell. */
function rangesOverlap(a, b) {
  const first = parseRange(a)
  const second = parseRange(b)
  return first.start.row <= second.end.row && second.start.row <= first.end.row
    && first.start.col <= second.end.col && second.start.col <= first.end.col
}

/**
 * The text a cell displays, as far as it can be known without a layout engine.
 * Shared strings are resolved through the workbook's string table.
 */
function cellDisplayText(cellXml, sharedStrings) {
  const type = attr(cellXml, 't') || 'n'
  const valueElement = extractElements(cellXml, 'v')[0]

  if (type === 's' && valueElement) {
    const index = Number(extractTextFromXml(valueElement.outerXml))
    return Number.isInteger(index) ? (sharedStrings[index] ?? '') : ''
  }

  const inline = extractElements(cellXml, 'is')[0]
  if (inline) return extractTextFromXml(inline.outerXml)
  if (valueElement) return extractTextFromXml(valueElement.outerXml)
  return ''
}

/**
 * High-performance XLSX processing engine for R7-Office.
 */
export class XlsxEngine {
  constructor(r7Adapter = null) {
    this.r7Adapter = r7Adapter || new R7Adapter()
  }

  /**
   * Inspect XLSX workbook structure, sheets, dimensions, formulas.
   * @param {string} filePath
   * @returns {Promise<object>}
   */
  async inspect(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const workbookXml = zip.getText('xl/workbook.xml')
    if (!workbookXml) throw new Error('Invalid XLSX: xl/workbook.xml not found')

    const sheets = []
    const sheetElements = extractElements(workbookXml, 'sheet')

    for (let i = 0; i < sheetElements.length; i++) {
      const s = sheetElements[i].outerXml
      const name = attr(s, 'name') || `Sheet${i + 1}`
      const sheetId = attr(s, 'sheetId') || `${i + 1}`
      const rId = attr(s, 'r:id') || `rId${i + 1}`

      const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
      let dimension = 'A1'
      let rowCount = 0
      let cellCount = 0

      if (zip.has(sheetPath)) {
        const sheetXml = zip.getText(sheetPath)
        const dimMatch = sheetXml.match(/<dimension\s+ref=["']([^"']+)["']/i)
        if (dimMatch) dimension = dimMatch[1]

        const rows = extractElements(sheetXml, 'row')
        rowCount = rows.length
        const cells = extractElements(sheetXml, 'c')
        cellCount = cells.length
      }

      sheets.push({
        index: i,
        name,
        sheetId,
        sheetPath,
        dimension,
        rowCount,
        cellCount
      })
    }

    return {
      type: 'xlsx',
      filePath,
      sheetsCount: sheets.length,
      sheets
    }
  }

  /**
   * Resolve a worksheet part path from the workbook's own part graph.
   *
   * The sheet order in `workbook.xml` is authoritative, and the r:id of each
   * `<sheet>` is resolved through `xl/_rels/workbook.xml.rels` — a worksheet's
   * file number need not match its position, which is exactly the case in
   * workbooks produced by other tools.
   *
   * @param {ZipArchive} zip
   * @param {{sheetIndex?: number, sheetName?: string}} target
   * @returns {string}
   */
  _resolveSheetPath(zip, target = {}) {
    const workbookXml = zip.getText('xl/workbook.xml')
    if (!workbookXml) {
      throw new Error('Invalid XLSX: xl/workbook.xml not found')
    }

    const sheets = extractElements(workbookXml, 'sheet')
    if (sheets.length === 0) {
      throw new Error('Invalid XLSX: the workbook declares no sheets')
    }

    let index = Number.isInteger(target.sheetIndex) ? target.sheetIndex : 0
    if (target.sheetName) {
      const wanted = target.sheetName.toLowerCase()
      const found = sheets.findIndex((s) => {
        const name = attr(s.outerXml, 'name')
        return name && name.toLowerCase() === wanted
      })
      if (found === -1) {
        const names = sheets.map((s) => attr(s.outerXml, 'name')).join(', ')
        throw new Error(`Sheet not found: "${target.sheetName}". Available sheets: ${names}`)
      }
      index = found
    }

    if (index < 0 || index >= sheets.length) {
      throw new Error(`Sheet index out of range: ${index} (the workbook has ${sheets.length} sheet(s))`)
    }

    const rId = attr(sheets[index].outerXml, 'r:id')
    const relsXml = zip.getText('xl/_rels/workbook.xml.rels')
    if (rId && relsXml) {
      const rel = extractElements(relsXml, 'Relationship')
        .find((r) => attr(r.outerXml, 'Id') === rId)
      const relTarget = rel ? attr(rel.outerXml, 'Target') : null
      if (relTarget) {
        const normalized = relTarget.replace(/^\/xl\//, '').replace(/^\.\//, '')
        const candidate = normalized.startsWith('xl/') ? normalized : `xl/${normalized}`
        if (zip.has(candidate)) return candidate
      }
    }

    // Fall back to the positional convention when the rels part is missing.
    const fallback = `xl/worksheets/sheet${index + 1}.xml`
    if (!zip.has(fallback)) {
      throw new Error(`Worksheet part not found for sheet index ${index}`)
    }
    return fallback
  }

  /**
   * Read cells from specific sheet and optional range (e.g. A1:D10).
   * @param {string} filePath
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async read(filePath, options = {}) {
    const {
      sheetIndex,
      sheetName = null,
      range = null,
      includeFormulas = false,
      includeStyles = false
    } = options
    const zip = await ZipArchive.fromFile(filePath)

    // Load shared strings
    const sharedStrings = this._loadSharedStrings(zip)

    const targetSheetPath = this._resolveSheetPath(zip, { sheetIndex, sheetName })

    const sheetXml = zip.getText(targetSheetPath)
    const styles = includeStyles ? this._loadStyles(zip) : null
    const rows = extractElements(sheetXml, 'row')
    const grid = {}
    let minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1

    for (const r of rows) {
      const rXml = r.outerXml
      const rowNum = parseInt(attr(rXml, 'r') || '1', 10) - 1
      const cells = extractElements(rXml, 'c')

      for (const c of cells) {
        const cXml = c.outerXml
        const cellRef = attr(cXml, 'r')
        if (!cellRef) continue

        const { col, row } = parseCellRef(cellRef)
        const cellType = attr(cXml, 't') || 'n'

        let val = ''
        const vElem = extractElements(cXml, 'v')[0]
        const fElem = extractElements(cXml, 'f')[0]
        const isElem = extractElements(cXml, 'is')[0]
        const tElem = extractElements(cXml, 't')[0]
        const formula = fElem ? extractTextFromXml(fElem.outerXml) : null

        if (cellType === 's' && vElem) {
          const sIdx = parseInt(extractTextFromXml(vElem.outerXml), 10)
          val = sharedStrings[sIdx] || ''
        } else if (cellType === 'inlineStr' || isElem || cellType === 'str') {
          if (isElem) {
            val = extractTextFromXml(isElem.outerXml)
          } else if (tElem) {
            val = extractTextFromXml(tElem.outerXml)
          } else if (vElem) {
            val = extractTextFromXml(vElem.outerXml)
          }
        } else if (vElem) {
          val = extractTextFromXml(vElem.outerXml)
          if (cellType === 'n' && !isNaN(Number(val))) {
            val = Number(val)
          }
        }

        minRow = Math.min(minRow, row)
        maxRow = Math.max(maxRow, row)
        minCol = Math.min(minCol, col)
        maxCol = Math.max(maxCol, col)

        if (!grid[row]) grid[row] = {}
        grid[row][col] = {
          ref: cellRef,
          value: val,
          formula,
          xfIndex: attr(cXml, 's') === null || attr(cXml, 's') === undefined
            ? null
            : Number(attr(cXml, 's'))
        }
      }
    }

    // Filter by range if specified (e.g. A1:C5)
    let startRow = 0, endRow = maxRow, startCol = 0, endCol = maxCol
    if (range) {
      const parts = range.split(':')
      const startRef = parseCellRef(parts[0])
      startRow = startRef.row
      startCol = startRef.col

      if (parts[1]) {
        const endRef = parseCellRef(parts[1])
        endRow = endRef.row
        endCol = endRef.col
      } else {
        endRow = startRow
        endCol = startCol
      }
    } else {
      startRow = minRow === Infinity ? 0 : minRow
      startCol = minCol === Infinity ? 0 : minCol
      endRow = maxRow === -1 ? 0 : maxRow
      endCol = maxCol === -1 ? 0 : maxCol
    }

    const dataMatrix = []
    const formulaMatrix = []
    const styleMatrix = []
    let formulaCount = 0

    for (let r = startRow; r <= endRow; r++) {
      const rowArr = []
      const formulaArr = []
      const styleArr = []
      for (let c = startCol; c <= endCol; c++) {
        const cellObj = grid[r]?.[c]
        rowArr.push(cellObj ? cellObj.value : '')
        formulaArr.push(cellObj?.formula ?? null)
        // Only a cell that exists can carry formatting. A cell that exists but
        // has no `s` attribute is cell format 0 — the default — which is a
        // different fact from "there is no cell here at all".
        styleArr.push(styles && cellObj ? styles.describeCellXf(cellObj.xfIndex ?? 0) : null)
        if (cellObj?.formula) formulaCount++
      }
      dataMatrix.push(rowArr)
      formulaMatrix.push(formulaArr)
      styleMatrix.push(styleArr)
    }

    const result = {
      sheet: targetSheetPath,
      sheetName: options.sheetName || undefined,
      range: range || `${indexToColLetter(startCol)}${startRow + 1}:${indexToColLetter(endCol)}${endRow + 1}`,
      rowCount: dataMatrix.length,
      colCount: dataMatrix[0]?.length || 0,
      data: dataMatrix
    }
    if (result.sheetName === undefined) delete result.sheetName

    // A formula cell only carries a cached <v> when a spreadsheet engine has
    // already calculated it. Exposing the formulas (and saying whether the
    // values are computed) is what lets a caller verify a formula was written
    // rather than misreading an empty string as a lost value.
    if (includeFormulas) {
      result.formulas = formulaMatrix
      result.formulaCount = formulaCount
      if (formulaCount > 0) {
        result.note = 'Formula cells are stored as written. Cached values are present only once a '
          + 'spreadsheet engine (for example R7-Office) has opened and recalculated the workbook.'
      }
    }

    // The style of every cell in the returned range, so a caller can see the
    // current formatting before deciding how to change it. Merged ranges, row
    // heights and column widths are sheet-level facts, reported once.
    if (includeStyles) {
      result.styles = styleMatrix
      result.merged = getMergedCells(sheetXml)
      result.rowHeights = getRowHeights(sheetXml)
      result.columnWidths = getColumnWidths(sheetXml)
      result.styleVocabulary = {
        note: 'Every value below is normalized; no raw OOXML is required to change it.',
        fontKeys: ['family', 'size', 'bold', 'italic', 'underline', 'strike', 'color'],
        fillKeys: ['color', 'patternType'],
        borderKeys: ['left', 'right', 'top', 'bottom'],
        alignmentKeys: ['horizontal', 'vertical', 'wrapText', 'textRotation'],
        numberFormatFields: {
          numberFormat: 'the declared format code, or the friendly name of a built-in id',
          numberFormatCode: 'the literal format code, or null when the id has no fixed code',
          numberFormatName: 'the friendly built-in name, or null for a custom format',
          numberFormatId: 'the numFmtId the cell actually refers to'
        },
        numberFormatExamples: ['integer', 'decimal', 'currency', 'percent', 'date', 'datetime', 'custom'],
        nullStyleMeans: 'The cell does not exist in the sheet. A cell that exists without an explicit '
          + 'style is reported with the default format (index 0), never as null.'
      }
    }

    return result
  }

  /**
   * Write values, dates or formulas into an XLSX sheet.
   *
   * Every update rewrites exactly one cell element in place: its payload and
   * type change, its style and its formula-related attributes survive, and the
   * cell is re-emitted at its ordered position in the row. A write therefore
   * never reorders a row, never unformats a cell it fills, and never touches
   * another part of the package.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {Array<{ref: string, value?: *, formula?: string, date?: boolean, numberFormat?: object}>} [options.cells]
   * @returns {Promise<{success: boolean, updatedCells: number, outputPath: string, sheet: string}>}
   */
  async write(filePath, options = {}) {
    const {
      outputPath = filePath,
      sheetIndex,
      sheetName = null,
      cells = [],
      matrix = null,
      startCell = 'A1'
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const targetSheetPath = this._resolveSheetPath(zip, { sheetIndex, sheetName })

    let sheetXml = zip.getText(targetSheetPath)
    if (sheetXml === null) throw new Error(`Worksheet part not found: ${targetSheetPath}`)

    // Ensure <sheetData> is expanded if self-closing
    if (sheetXml.includes('<sheetData/>')) {
      sheetXml = sheetXml.replace('<sheetData/>', '<sheetData></sheetData>')
    }

    const updates = new Map()

    if (matrix && Array.isArray(matrix)) {
      const start = parseCellRef(startCell)
      for (let r = 0; r < matrix.length; r++) {
        const rowData = matrix[r]
        for (let c = 0; c < rowData.length; c++) {
          const colLetter = indexToColLetter(start.col + c)
          const ref = `${colLetter}${start.rowNum + r}`
          updates.set(ref, { ref, value: rowData[c] })
        }
      }
    }

    for (const item of cells) {
      if (item && item.ref) {
        const ref = String(item.ref).toUpperCase()
        updates.set(ref, { ...item, ref })
      }
    }

    // Only a real date or an explicit number format needs the stylesheet;
    // a plain value write stays a single-part edit.
    const needsStyles = [...updates.values()].some((u) => u.date === true || u.numberFormat)
    const styles = needsStyles ? this._loadStyles(zip) : null
    const stylesExisted = zip.has('xl/styles.xml')
    const xfCache = new Map()

    for (const [ref, update] of updates.entries()) {
      const base = getCellStyle(sheetXml, ref) ?? 0
      const formula = update.formula ? String(update.formula).replace(/^=/, '') : null
      let value = update.value
      let numberFormat = update.numberFormat || null

      // A date is stored as the serial number a spreadsheet counts days with —
      // never as text — so it stays sortable and can be reformatted later.
      if (update.date === true) {
        value = dateToSerial(value)
        if (!numberFormat) {
          numberFormat = { type: hasTimePart(update.value) ? 'datetime' : 'date' }
        }
      }

      const isNumber = typeof value === 'number' && Number.isFinite(value)
      let type = null
      let content = ''

      if (formula) {
        content += `<f>${escapeXml(formula)}</f>`
        if (isNumber) {
          content += `<v>${value}</v>`
        } else if (value !== undefined && value !== null && value !== '') {
          type = 'str'
          content += `<v>${escapeXml(String(value))}</v>`
        }
      } else if (isNumber) {
        content += `<v>${value}</v>`
      } else if (value !== undefined && value !== null && value !== '') {
        type = 'inlineStr'
        content += `<is><t>${escapeXml(String(value))}</t></is>`
      } else {
        type = 'inlineStr'
      }

      let xfIndex
      if (styles) {
        const key = `${base}|${JSON.stringify(numberFormat || null)}`
        if (!xfCache.has(key)) {
          xfCache.set(key, styles.ensureCellXf({ base, numberFormat }))
        }
        xfIndex = xfCache.get(key)
      }

      sheetXml = setCellContent(sheetXml, ref, { content, type, xfIndex })
    }

    zip.setText(targetSheetPath, sheetXml)
    if (styles?.dirty) this._writeStyles(zip, styles, stylesExisted)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    await zip.save(outputPath)
    return { success: true, updatedCells: updates.size, outputPath, sheet: targetSheetPath }
  }

  /**
   * Apply formatting to a cell or a range of an XLSX worksheet.
   *
   * The whole operation is "based on what is already there": each cell keeps
   * the parts of its format the caller did not mention (including its borders,
   * fill and number format), and the stylesheet keeps every table entry at the
   * index the rest of the workbook already refers to.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {string} options.range - required; "B2" or "A1:D10"
   * @param {object} [options.font] - { family, size, bold, italic, underline, strike, color }
   * @param {object} [options.fill] - { color }
   * @param {object} [options.border] - { all|top|bottom|left|right } each 'thin' | { style, color }
   * @param {object} [options.alignment] - { horizontal, vertical, wrapText, textRotation }
   * @param {object} [options.numberFormat] - { type, decimals, symbol, thousands, code }
   * @param {boolean} [options.merge]
   * @param {boolean} [options.unmerge]
   * @param {object} [options.columnWidth] - { width } or { auto: true }
   * @param {number} [options.rowHeight] - points
   * @returns {Promise<object>}
   */
  async format(filePath, options = {}) {
    const {
      outputPath = filePath,
      sheetIndex,
      sheetName = null,
      range,
      font = null,
      fill = null,
      border = null,
      alignment = null,
      numberFormat = null,
      merge = false,
      unmerge = false,
      columnWidth = null,
      rowHeight = null
    } = options

    if (typeof range !== 'string' || range.trim() === '') {
      throw new Error('format requires a range (a cell such as "B2" or a range such as "A1:D10")')
    }

    const normalizedRange = normalizeRange(range)
    const { start, end } = parseRange(normalizedRange)

    const zip = await ZipArchive.fromFile(filePath)
    const targetSheetPath = this._resolveSheetPath(zip, { sheetIndex, sheetName })
    let sheetXml = zip.getText(targetSheetPath)
    if (sheetXml === null) throw new Error(`Worksheet part not found: ${targetSheetPath}`)

    const stylesExisted = zip.has('xl/styles.xml')
    const styles = this._loadStyles(zip)
    const borderSpec = border ? expandBorderSpec(border) : null
    const hasStyleRequest = Boolean(font || fill || borderSpec || alignment || numberFormat)

    const refs = expandRange(normalizedRange)
    // Every cell format is derived from the style the cell had *before* this
    // call, so the base index is read up front rather than while the sheet is
    // being rewritten.
    const bases = new Map(refs.map((ref) => [ref, getCellStyle(sheetXml, ref) ?? 0]))
    const xfCache = new Map()

    if (hasStyleRequest) {
      // Cells that already share a base format share the resulting entry, so a
      // thousand-cell range adds one cellXfs entry per distinct input format.
      for (const ref of refs) {
        const base = bases.get(ref)
        if (!xfCache.has(base)) {
          xfCache.set(base, styles.ensureCellXf({
            base,
            font,
            fill,
            border: borderSpec,
            alignment,
            numberFormat
          }))
        }
        sheetXml = setCellStyle(sheetXml, ref, xfCache.get(base))
      }
    }

    if (unmerge) {
      // "Remove the merge covering this range" also has to cover the case where
      // the caller names a single cell inside a merged block.
      const covering = getMergedCells(sheetXml).filter((r) => rangesOverlap(r, normalizedRange))
      if (covering.length > 0) sheetXml = setMergedCells(sheetXml, { remove: covering })
    }
    if (merge) sheetXml = setMergedCells(sheetXml, { add: [normalizedRange] })

    if (columnWidth) {
      const explicit = Number(columnWidth.width)
      if (columnWidth.auto === true) {
        const sharedStrings = this._loadSharedStrings(zip)
        for (let col = start.col; col <= end.col; col++) {
          const width = this._autoColumnWidth(sheetXml, col, sharedStrings)
          sheetXml = setColumnWidth(sheetXml, { min: col + 1, max: col + 1, width })
        }
      } else if (Number.isFinite(explicit) && explicit > 0) {
        sheetXml = setColumnWidth(sheetXml, { min: start.col + 1, max: end.col + 1, width: explicit })
      } else {
        throw new Error('columnWidth requires a positive "width", or { auto: true }')
      }
    }

    if (rowHeight !== null && rowHeight !== undefined) {
      for (let row = start.row; row <= end.row; row++) {
        sheetXml = setRowHeight(sheetXml, row + 1, rowHeight)
      }
    }

    zip.setText(targetSheetPath, sheetXml)
    if (styles.dirty) this._writeStyles(zip, styles, stylesExisted)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    return {
      success: true,
      outputPath,
      sheet: targetSheetPath,
      range: normalizedRange,
      cellsFormatted: hasStyleRequest ? refs.length : 0,
      merges: getMergedCells(sheetXml),
      columnWidths: getColumnWidths(sheetXml),
      rowHeights: getRowHeights(sheetXml)
    }
  }

  /**
   * The stylesheet of a workbook, or a fresh one when the package has none.
   * @param {ZipArchive} zip
   * @returns {Stylesheet}
   */
  _loadStyles(zip) {
    const xml = zip.getText('xl/styles.xml')
    if (!xml) return new Stylesheet()
    try {
      return new Stylesheet(xml)
    } catch (err) {
      throw new Error(`Invalid XLSX: cannot read xl/styles.xml: ${err.message}`)
    }
  }

  /**
   * Store a changed stylesheet, registering the package part when the workbook
   * did not have one (a workbook built without the R7 template).
   */
  _writeStyles(zip, styles, existed) {
    if (!existed) this._registerStylesPart(zip)
    zip.setText('xl/styles.xml', styles.toXml())
  }

  /** Declare `xl/styles.xml` in [Content_Types].xml and the workbook rels. */
  _registerStylesPart(zip) {
    const contentTypesPath = '[Content_Types].xml'
    const contentTypes = zip.getText(contentTypesPath)
    if (contentTypes && !contentTypes.includes('PartName="/xl/styles.xml"')) {
      zip.setText(contentTypesPath, contentTypes.replace(
        '</Types>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
      ))
    }

    const relsPath = 'xl/_rels/workbook.xml.rels'
    const rels = zip.getText(relsPath)
    if (rels && !rels.includes('relationships/styles')) {
      const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
      const relId = `rId${ids.length > 0 ? Math.max(...ids) + 1 : 1}`
      zip.setText(relsPath, rels.replace(
        '</Relationships>',
        `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`
      ))
    }
  }

  /**
   * A width that fits every value written in a column.
   *
   * The estimate is heuristic — the real rendered width depends on the font a
   * spreadsheet engine resolves at layout time — so it is deliberately wide
   * rather than exact.
   */
  _autoColumnWidth(sheetXml, colIndex, sharedStrings) {
    const values = []
    for (const row of extractElements(sheetXml, 'row')) {
      for (const cell of extractElements(row.outerXml, 'c')) {
        const ref = attr(cell.outerXml, 'r')
        if (!ref) continue
        if (parseCellRef(ref).col !== colIndex) continue
        const text = cellDisplayText(cell.outerXml, sharedStrings)
        if (text !== '') values.push({ text })
      }
    }
    return estimateColumnWidth(values)
  }

  /**
   * Add a worksheet to an existing workbook.
   *
   * Registers every required package part: the worksheet itself, its content
   * type override, the workbook `<sheet>` entry and the workbook relationship.
   * Existing worksheets are never rewritten.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {string} [options.name] - Worksheet name (defaults to "ЛистN").
   * @param {number} [options.index] - Insert position; appended when omitted.
   * @param {Array<Array<*>>} [options.data] - Optional 2-D data to write.
   * @param {string} [options.outputPath]
   * @returns {Promise<{success: boolean, name: string, sheetIndex: number, outputPath: string}>}
   */
  async addSheet(filePath, options = {}) {
    const {
      name = null,
      index = null,
      data = null,
      outputPath = filePath
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const workbookXml = zip.getText('xl/workbook.xml')
    if (!workbookXml) throw new Error('Invalid XLSX: xl/workbook.xml not found')

    const existingSheets = extractElements(workbookXml, 'sheet')
    const usedNames = existingSheets
      .map((s) => attr(s.outerXml, 'name'))
      .filter(Boolean)
    const usedSheetIds = existingSheets
      .map((s) => Number(attr(s.outerXml, 'sheetId')))
      .filter((n) => Number.isFinite(n))

    const sheetName = this._uniqueSheetName(name || `Лист${usedNames.length + 1}`, usedNames)

    // A worksheet part number must be free; the position in the tab order is a
    // separate thing, carried by the order of <sheet> in workbook.xml.
    const partNumbers = zip.list()
      .map((entry) => {
        const m = entry.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)
        return m ? Number(m[1]) : 0
      })
      .filter(Boolean)
    const partNumber = partNumbers.length > 0 ? Math.max(...partNumbers) + 1 : 1
    const sheetPath = `xl/worksheets/sheet${partNumber}.xml`
    const sheetId = usedSheetIds.length > 0 ? Math.max(...usedSheetIds) + 1 : 1

    // Relationship id for the workbook -> worksheet edge.
    const relsPath = 'xl/_rels/workbook.xml.rels'
    let relsXml = zip.getText(relsPath)
    if (!relsXml) {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    }
    const relIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
    const relId = `rId${relIds.length > 0 ? Math.max(...relIds) + 1 : 1}`
    relsXml = relsXml.replace(
      '</Relationships>',
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${partNumber}.xml"/></Relationships>`
    )

    // Content type override for the new part.
    const contentTypesPath = '[Content_Types].xml'
    let contentTypes = zip.getText(contentTypesPath)
    if (!contentTypes) throw new Error('Invalid XLSX: [Content_Types].xml not found')
    if (!contentTypes.includes(`/xl/worksheets/sheet${partNumber}.xml`)) {
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override PartName="/xl/worksheets/sheet${partNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
      )
    }

    // The new sheet reuses the shell of an existing one (namespaces, views,
    // formatting defaults) with an empty sheetData.
    const basePath = this._resolveSheetPath(zip, { sheetIndex: 0 })
    const baseXml = zip.getText(basePath) || ''
    const blankSheetXml = baseXml.includes('<sheetData')
      ? baseXml.replace(/<sheetData[\s\S]*?<\/sheetData>|<sheetData\s*\/>/, '<sheetData/>')
      : this._blankWorksheetXml()

    const newSheetEntry = `<sheet name="${escapeXml(sheetName)}" sheetId="${sheetId}" r:id="${relId}"/>`
    let updatedWorkbook = workbookXml
    if (index !== null && Number.isInteger(index) && index >= 0 && index < existingSheets.length) {
      updatedWorkbook = updatedWorkbook.replace(
        existingSheets[index].outerXml,
        `${newSheetEntry}${existingSheets[index].outerXml}`
      )
    } else {
      updatedWorkbook = updatedWorkbook.replace('</sheets>', `${newSheetEntry}</sheets>`)
    }

    zip.setBuffer(sheetPath, Buffer.from(blankSheetXml, 'utf8'))
    zip.setText(relsPath, relsXml)
    zip.setText(contentTypesPath, contentTypes)
    zip.setText('xl/workbook.xml', updatedWorkbook)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    const sheetIndex = index !== null && Number.isInteger(index) && index >= 0
      ? Math.min(index, existingSheets.length)
      : existingSheets.length

    if (data && Array.isArray(data) && data.length > 0) {
      await this.write(outputPath, { sheetName, matrix: data, startCell: 'A1' })
    }

    return { success: true, name: sheetName, sheetIndex, outputPath }
  }

  /** Keep worksheet names unique, as Excel and R7 require. */
  _uniqueSheetName(desired, usedNames) {
    const taken = new Set(usedNames.map((n) => n.toLowerCase()))
    if (!taken.has(desired.toLowerCase())) return desired
    let n = 2
    while (taken.has(`${desired}${n}`.toLowerCase())) n++
    return `${desired}${n}`
  }

  _blankWorksheetXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<dimension ref="A1"/><sheetData/></worksheet>'
  }

  /**
   * Create a new XLSX workbook from template or clean structure.
   * @param {string} outputPath
   * @param {object} [options]
   * @returns {Promise<{success: boolean, path: string}>}
   */
  async create(outputPath, options = {}) {
    const { sheets = [{ name: 'Лист1', data: [] }], overwrite = false } = options

    if (fs.existsSync(outputPath) && overwrite !== true) {
      throw new Error(
        `Refusing to overwrite: ${outputPath} already exists. Pass overwrite: true to replace it.`
      )
    }

    const tplPath = await this.r7Adapter.getTemplatePath('xlsx')
    let zip

    if (tplPath && fs.existsSync(tplPath)) {
      zip = await ZipArchive.fromFile(tplPath)
    } else {
      zip = this._createFallbackBlankXlsx()
    }

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    // The template ships exactly one worksheet; rename it to the caller's
    // first sheet so every requested name is honoured.
    const requested = Array.isArray(sheets) && sheets.length > 0
      ? sheets
      : [{ name: 'Лист1', data: [] }]
    const firstName = requested[0].name || 'Лист1'
    const workbookXml = zip.getText('xl/workbook.xml') || ''
    const firstSheet = extractElements(workbookXml, 'sheet')[0]
    if (firstSheet) {
      const original = firstSheet.outerXml
      const renamed = original.replace(/\sname="[^"]*"/, ` name="${escapeXml(firstName)}"`)
      zip.setText('xl/workbook.xml', workbookXml.replace(original, renamed))
    }

    await zip.save(outputPath)

    // Fill sheet 1, then add and fill the rest.
    if (requested[0].data?.length > 0) {
      await this.write(outputPath, { sheetIndex: 0, matrix: requested[0].data, startCell: 'A1' })
    }

    for (let i = 1; i < requested.length; i++) {
      await this.addSheet(outputPath, {
        name: requested[i].name || `Лист${i + 1}`,
        data: requested[i].data || null
      })
    }

    return { success: true, path: outputPath }
  }

  /**
   * Validate XLSX integrity.
   * @param {string} filePath
   * @returns {Promise<{valid: boolean, errors: string[], details: object}>}
   */
  async validate(filePath) {
    const errors = []
    let zip
    try {
      zip = await ZipArchive.fromFile(filePath)
    } catch (err) {
      return { valid: false, errors: [`ZIP corruption: ${err.message}`], details: {} }
    }

    if (!zip.has('[Content_Types].xml')) errors.push('Missing [Content_Types].xml')
    if (!zip.has('xl/workbook.xml')) errors.push('Missing xl/workbook.xml')
    if (!zip.has('xl/worksheets/sheet1.xml')) errors.push('Missing xl/worksheets/sheet1.xml')

    return {
      valid: errors.length === 0,
      errors,
      details: {
        fileSize: fs.statSync(filePath).size,
        entries: zip.list().length
      }
    }
  }

  _loadSharedStrings(zip) {
    const strings = []
    const sstXml = zip.getText('xl/sharedStrings.xml')
    if (!sstXml) return strings

    const siElements = extractElements(sstXml, 'si')
    for (const si of siElements) {
      strings.push(extractTextFromXml(si.outerXml))
    }
    return strings
  }

  _createFallbackBlankXlsx() {
    const zip = new ZipArchive()
    zip.setText('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`)

    zip.setText('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)

    zip.setText('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Лист1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`)

    zip.setText('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`)

    zip.setText('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>`)

    return zip
  }
}
