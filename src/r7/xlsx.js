import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements, getAttribute } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'

/**
 * Helper to convert column letter (A, B, AA) to 0-based index and vice versa.
 */
export function colLetterToIndex(colStr) {
  let index = 0
  for (let i = 0; i < colStr.length; i++) {
    index = index * 26 + (colStr.charCodeAt(i) - 64)
  }
  return index - 1
}

export function indexToColLetter(index) {
  let colStr = ''
  let n = index + 1
  while (n > 0) {
    let rem = (n - 1) % 26
    colStr = String.fromCharCode(65 + rem) + colStr
    n = Math.floor((n - 1) / 26)
  }
  return colStr
}

export function parseCellRef(cellRef) {
  const match = cellRef.match(/^([A-Za-z]+)([0-9]+)$/)
  if (!match) throw new Error(`Invalid cell reference: ${cellRef}`)
  return {
    colStr: match[1].toUpperCase(),
    col: colLetterToIndex(match[1].toUpperCase()),
    row: parseInt(match[2], 10) - 1,
    rowNum: parseInt(match[2], 10)
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
   * Read cells from specific sheet and optional range (e.g. A1:D10).
   * @param {string} filePath
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async read(filePath, options = {}) {
    const { sheetIndex = 0, sheetName = null, range = null } = options
    const zip = await ZipArchive.fromFile(filePath)

    // Load shared strings
    const sharedStrings = this._loadSharedStrings(zip)

    // Determine target sheet path
    let targetSheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`
    const workbookXml = zip.getText('xl/workbook.xml')

    if (sheetName && workbookXml) {
      const sheetElements = extractElements(workbookXml, 'sheet')
      for (let i = 0; i < sheetElements.length; i++) {
        const sName = getAttribute(sheetElements[i].outerXml, 'name')
        if (sName && sName.toLowerCase() === sheetName.toLowerCase()) {
          targetSheetPath = `xl/worksheets/sheet${i + 1}.xml`
          break
        }
      }
    }

    if (!zip.has(targetSheetPath)) {
      throw new Error(`Sheet not found: ${targetSheetPath}`)
    }

    const sheetXml = zip.getText(targetSheetPath)
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
          formula
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
    for (let r = startRow; r <= endRow; r++) {
      const rowArr = []
      for (let c = startCol; c <= endCol; c++) {
        const cellObj = grid[r]?.[c]
        rowArr.push(cellObj ? cellObj.value : '')
      }
      dataMatrix.push(rowArr)
    }

    return {
      sheet: targetSheetPath,
      range: range || `${indexToColLetter(startCol)}${startRow + 1}:${indexToColLetter(endCol)}${endRow + 1}`,
      rowCount: dataMatrix.length,
      colCount: dataMatrix[0]?.length || 0,
      data: dataMatrix
    }
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
      sheetIndex = 0,
      cells = [],
      matrix = null,
      startCell = 'A1'
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const targetSheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`

    if (!zip.has(targetSheetPath)) {
      throw new Error(`Sheet not found: ${targetSheetPath}`)
    }

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
   * Create a new XLSX workbook from template or clean structure.
   * @param {string} outputPath
   * @param {object} [options]
   * @returns {Promise<{success: boolean, path: string}>}
   */
  async create(outputPath, options = {}) {
    const { sheets = [{ name: 'Лист1', data: [] }] } = options

    const tplPath = await this.r7Adapter.getTemplatePath('xlsx')
    let zip

    if (tplPath && fs.existsSync(tplPath)) {
      zip = await ZipArchive.fromFile(tplPath)
    } else {
      zip = this._createFallbackBlankXlsx()
    }

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    // Write initial data if provided
    if (sheets[0]?.data?.length > 0) {
      await this.write(outputPath, { matrix: sheets[0].data, startCell: 'A1' })
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
