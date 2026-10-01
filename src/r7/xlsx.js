import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements, getAttribute } from '../shared/xml.js'
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
  buildRef,
  getCellStyle,
  setCellStyle,
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
      const name = getAttribute(s, 'name') || `Sheet${i + 1}`
      const sheetId = getAttribute(s, 'sheetId') || `${i + 1}`
      const rId = getAttribute(s, 'r:id') || `rId${i + 1}`

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
        const name = getAttribute(s.outerXml, 'name')
        return name && name.toLowerCase() === wanted
      })
      if (found === -1) {
        const names = sheets.map((s) => getAttribute(s.outerXml, 'name')).join(', ')
        throw new Error(`Sheet not found: "${target.sheetName}". Available sheets: ${names}`)
      }
      index = found
    }

    if (index < 0 || index >= sheets.length) {
      throw new Error(`Sheet index out of range: ${index} (the workbook has ${sheets.length} sheet(s))`)
    }

    const rId = getAttribute(sheets[index].outerXml, 'r:id')
    const relsXml = zip.getText('xl/_rels/workbook.xml.rels')
    if (rId && relsXml) {
      const rel = extractElements(relsXml, 'Relationship')
        .find((r) => getAttribute(r.outerXml, 'Id') === rId)
      const relTarget = rel ? getAttribute(rel.outerXml, 'Target') : null
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
      const rowNum = parseInt(getAttribute(rXml, 'r') || '1', 10) - 1
      const cells = extractElements(rXml, 'c')

      for (const c of cells) {
        const cXml = c.outerXml
        const cellRef = getAttribute(cXml, 'r')
        if (!cellRef) continue

        const { col, row } = parseCellRef(cellRef)
        const cellType = getAttribute(cXml, 't') || 'n'

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
          xfIndex: getAttribute(cXml, 's') === null || getAttribute(cXml, 's') === undefined
            ? null
            : Number(getAttribute(cXml, 's'))
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
        styleArr.push(styles ? styles.describeCellXf(cellObj?.xfIndex ?? 0) : null)
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
        alignmentKeys: ['horizontal', 'vertical', 'wrapText'],
        numberFormatExamples: ['integer', 'decimal', 'currency', 'percent', 'date', 'datetime', 'custom']
      }
    }

    return result
  }

  /**
   * Write values or formulas into XLSX sheet.
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
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
        updates.set(item.ref.toUpperCase(), item)
      }
    }

    // Apply updates directly into sheet XML
    for (const [ref, update] of updates.entries()) {
      const val = update.value
      const formula = update.formula
      const isNum = typeof val === 'number'
      const typeAttr = isNum ? '' : ' t="inlineStr"'

      let valContent = ''
      if (formula) {
        valContent += `<f>${escapeXml(formula.replace(/^=/, ''))}</f>`
      }
      if (isNum) {
        valContent += `<v>${val}</v>`
      } else if (val !== undefined && val !== null && val !== '') {
        valContent += `<is><t>${escapeXml(String(val))}</t></is>`
      }

      const newCellXml = `<c r="${ref}"${typeAttr}>${valContent}</c>`

      const cellRegex = new RegExp(`<c\\s+r=["']${ref}["'][^>]*>([\\s\\S]*?<\\/c>|\\/>)`, 'i')
      if (cellRegex.test(sheetXml)) {
        sheetXml = sheetXml.replace(cellRegex, newCellXml)
      } else {
        const { rowNum } = parseCellRef(ref)
        const rowRegex = new RegExp(`<row\\s+[^>]*r=["']${rowNum}["'][^>]*>([\\s\\S]*?)<\\/row>`, 'i')
        const rowMatch = sheetXml.match(rowRegex)

        if (rowMatch) {
          const updatedRow = rowMatch[0].replace('</row>', `${newCellXml}</row>`)
          sheetXml = sheetXml.replace(rowMatch[0], updatedRow)
        } else {
          const newRowXml = `<row r="${rowNum}">${newCellXml}</row>`
          sheetXml = sheetXml.replace('</sheetData>', `${newRowXml}</sheetData>`)
        }
      }
    }

    zip.setText(targetSheetPath, sheetXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    await zip.save(outputPath)
    return { success: true, updatedCells: updates.size, outputPath }
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
      .map((s) => getAttribute(s.outerXml, 'name'))
      .filter(Boolean)
    const usedSheetIds = existingSheets
      .map((s) => Number(getAttribute(s.outerXml, 'sheetId')))
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
