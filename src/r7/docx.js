import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements, getAttribute } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'

/**
 * High-performance, style-preserving DOCX processing engine for R7-Office.
 */
export class DocxEngine {
  constructor(r7Adapter = null) {
    this.r7Adapter = r7Adapter || new R7Adapter()
  }

  /**
   * Inspect document structure, outline, headings, paragraph count, and tables.
   * @param {string} filePath
   * @returns {Promise<object>}
   */
  async inspect(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) {
      throw new Error('Invalid DOCX: word/document.xml not found')
    }

    const paragraphs = []
    const headings = []
    const tables = []

    // Extract all paragraphs
    const pElements = extractElements(docXml, 'w:p')
    for (let i = 0; i < pElements.length; i++) {
      const p = pElements[i]
      const text = extractTextFromXml(p.outerXml).trim()
      const styleMatch = p.outerXml.match(/<w:pStyle\s+w:val=["']([^"']+)["']/i)
      const style = styleMatch ? styleMatch[1] : 'Normal'

      const isHeading = /heading\s*([0-9]+)/i.test(style) || /заголовок\s*([0-9]+)/i.test(style) || /heading/i.test(style)
      const headingLevel = isHeading ? parseInt(style.match(/[0-9]+/)?.[0] || '1', 10) : null

      const info = {
        index: i,
        style,
        isHeading,
        headingLevel,
        textLength: text.length,
        preview: text.substring(0, 100) + (text.length > 100 ? '...' : '')
      }
      paragraphs.push(info)

      if (isHeading && text) {
        headings.push({
          paragraphIndex: i,
          level: headingLevel,
          text
        })
      }
    }

    // Extract tables
    const tblElements = extractElements(docXml, 'w:tbl')
    for (let t = 0; t < tblElements.length; t++) {
      const tbl = tblElements[t]
      const rows = extractElements(tbl.outerXml, 'w:tr')
      const rowCount = rows.length
      let colCount = 0

      if (rowCount > 0) {
        const firstRowCells = extractElements(rows[0].outerXml, 'w:tc')
        colCount = firstRowCells.length
      }

      tables.push({
        tableIndex: t,
        rows: rowCount,
        cols: colCount
      })
    }

    // Read metadata if available
    const coreXml = zip.getText('docProps/core.xml')
    let metadata = {}
    if (coreXml) {
      metadata = {
        title: extractTextFromXml(extractElements(coreXml, 'dc:title')[0]?.outerXml || ''),
        creator: extractTextFromXml(extractElements(coreXml, 'dc:creator')[0]?.outerXml || ''),
        lastModifiedBy: extractTextFromXml(extractElements(coreXml, 'cp:lastModifiedBy')[0]?.outerXml || ''),
        revision: extractTextFromXml(extractElements(coreXml, 'cp:revision')[0]?.outerXml || '')
      }
    }

    return {
      type: 'docx',
      filePath,
      paragraphsCount: paragraphs.length,
      headingsCount: headings.length,
      tablesCount: tables.length,
      headings,
      tables,
      metadata,
      outline: paragraphs.slice(0, 50)
    }
  }

  /**
   * Read structured text, paragraphs, or tables from DOCX.
   * @param {string} filePath
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async read(filePath, options = {}) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) {
      throw new Error('Invalid DOCX: word/document.xml not found')
    }

    const { format = 'structured', fromParagraph = 0, count = 100, query = null } = options

    const pElements = extractElements(docXml, 'w:p')
    const paragraphs = []

    for (let i = 0; i < pElements.length; i++) {
      const p = pElements[i]
      const text = extractTextFromXml(p.outerXml)
      const styleMatch = p.outerXml.match(/<w:pStyle\s+w:val=["']([^"']+)["']/i)
      const style = styleMatch ? styleMatch[1] : 'Normal'

      if (query && !text.toLowerCase().includes(query.toLowerCase())) {
        continue
      }

      paragraphs.push({
        index: i,
        style,
        text
      })
    }

    const sliced = paragraphs.slice(fromParagraph, fromParagraph + count)

    if (format === 'markdown' || format === 'text') {
      const textContent = sliced.map(p => {
        if (/heading\s*1/i.test(p.style)) return `# ${p.text}`
        if (/heading\s*2/i.test(p.style)) return `## ${p.text}`
        if (/heading\s*3/i.test(p.style)) return `### ${p.text}`
        return p.text
      }).filter(t => t.trim().length > 0).join('\n\n')

      return {
        format,
        content: textContent,
        totalParagraphs: paragraphs.length,
        returned: sliced.length
      }
    }

    return {
      format: 'structured',
      totalParagraphs: paragraphs.length,
      paragraphs: sliced
    }
  }

  /**
   * Create a new DOCX document.
   * @param {string} outputPath
   * @param {object} [options]
   * @returns {Promise<{success: boolean, path: string}>}
   */
  async create(outputPath, options = {}) {
    const { title, paragraphs = [], tables = [], overwrite = false } = options

    if (fs.existsSync(outputPath) && overwrite !== true) {
      throw new Error(
        `Refusing to overwrite: ${outputPath} already exists. Pass overwrite: true to replace it.`
      )
    }

    // Try using native R7 template
    const tplPath = await this.r7Adapter.getTemplatePath('docx')
    let zip

    if (tplPath && fs.existsSync(tplPath)) {
      zip = await ZipArchive.fromFile(tplPath)
    } else {
      zip = this._createFallbackBlankDocx()
    }

    let docXml = zip.getText('word/document.xml') || ''

    // Build document body elements
    const bodyParts = []

    if (title) {
      bodyParts.push(
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>${escapeXml(title)}</w:t></w:r></w:p>`
      )
    }

    for (const p of paragraphs) {
      if (typeof p === 'string') {
        bodyParts.push(`<w:p><w:r><w:t>${escapeXml(p)}</w:t></w:r></w:p>`)
      } else if (p && typeof p === 'object') {
        const text = p.text || ''
        const style = p.style || 'Normal'
        const bold = p.bold ? '<w:b/>' : ''
        const italic = p.italic ? '<w:i/>' : ''
        const align = p.align ? `<w:jc w:val="${p.align}"/>` : ''
        const styleXml = style !== 'Normal' ? `<w:pStyle w:val="${style}"/>` : ''

        bodyParts.push(
          `<w:p><w:pPr>${styleXml}${align}</w:pPr>` +
          `<w:r><w:rPr>${bold}${italic}</w:rPr><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
        )
      }
    }

    // Build tables if requested
    for (const tbl of tables) {
      const rows = tbl.rows || []
      const tblXml = this._buildTableXml(rows)
      bodyParts.push(tblXml)
    }

    // Replace initial empty paragraphs in template
    if (bodyParts.length > 0) {
      const joined = bodyParts.join('\n')
      // Clean out initial empty <w:p> from template
      docXml = docXml.replace(/<w:body>\s*<w:p><w:r><\/w:r><w:r><\/w:r><\/w:p>/, '<w:body>')
      docXml = docXml.replace(/<w:body>\s*<w:p\/>/, '<w:body>')

      if (docXml.includes('<w:sectPr')) {
        docXml = docXml.replace('<w:sectPr', `${joined}\n<w:sectPr`)
      } else if (docXml.includes('</w:body>')) {
        docXml = docXml.replace('</w:body>', `${joined}\n</w:body>`)
      } else {
        docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${joined}<w:sectPr/></w:body></w:document>`
      }
      zip.setText('word/document.xml', docXml)
    }

    // Ensure output directory exists
    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true })
    }

    await zip.save(outputPath)
    return { success: true, path: outputPath }
  }

  /**
   * Search and replace text in DOCX while preserving formatting and styles.
   * @param {string} filePath
   * @param {string|RegExp} search
   * @param {string} replace
   * @param {object} [options]
   * @returns {Promise<{success: boolean, matchesCount: number, outputPath: string}>}
   */
  async replaceText(filePath, search, replace, options = {}) {
    const { outputPath = filePath, matchCase = true, replaceAll = true } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) {
      throw new Error('Invalid DOCX: word/document.xml not found')
    }

    let matchesCount = 0

    // Process paragraph by paragraph to preserve run formatting
    const pElements = extractElements(docXml, 'w:p')

    for (const p of pElements) {
      const pXml = p.outerXml
      const tElements = extractElements(pXml, 'w:t')

      if (tElements.length === 0) continue

      // Combine text across runs in this paragraph
      const combinedText = tElements.map(t => unescapeXml(t.innerXml)).join('')

      const isSearchMatch = typeof search === 'string'
        ? (matchCase ? combinedText.includes(search) : combinedText.toLowerCase().includes(search.toLowerCase()))
        : search.test(combinedText)

      if (!isSearchMatch) continue

      // Strategy 1: Direct single-run replacement if pattern fits inside one run
      let replacedPXml = pXml
      let paragraphMatched = false

      for (const t of tElements) {
        const runText = unescapeXml(t.innerXml)
        let newRunText = runText

        if (typeof search === 'string') {
          if (runText.includes(search)) {
            newRunText = replaceAll ? runText.replaceAll(search, replace) : runText.replace(search, replace)
            matchesCount++
            paragraphMatched = true
          } else if (!matchCase && runText.toLowerCase().includes(search.toLowerCase())) {
            const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), replaceAll ? 'gi' : 'i')
            newRunText = runText.replace(regex, replace)
            matchesCount++
            paragraphMatched = true
          }
        } else {
          if (search.test(runText)) {
            newRunText = runText.replace(search, replace)
            matchesCount++
            paragraphMatched = true
          }
        }

        if (newRunText !== runText) {
          const oldTTag = t.outerXml
          const newTTag = `<w:t xml:space="preserve">${escapeXml(newRunText)}</w:t>`
          replacedPXml = replacedPXml.replace(oldTTag, newTTag)
        }
      }

      // Strategy 2: Cross-run replacement (when search spans multiple <w:t> runs in paragraph)
      if (!paragraphMatched) {
        let newCombined = combinedText
        if (typeof search === 'string') {
          const regex = new RegExp(search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), (replaceAll ? 'g' : '') + (matchCase ? '' : 'i'))
          newCombined = combinedText.replace(regex, replace)
        } else {
          newCombined = combinedText.replace(search, replace)
        }

        if (newCombined !== combinedText) {
          matchesCount++
          // Put entire replaced text in the first run and empty the rest
          let firstUpdated = false
          for (const t of tElements) {
            const oldTTag = t.outerXml
            if (!firstUpdated) {
              const newTTag = `<w:t xml:space="preserve">${escapeXml(newCombined)}</w:t>`
              replacedPXml = replacedPXml.replace(oldTTag, newTTag)
              firstUpdated = true
            } else {
              replacedPXml = replacedPXml.replace(oldTTag, `<w:t></w:t>`)
            }
          }
        }
      }

      if (replacedPXml !== pXml) {
        docXml = docXml.replace(pXml, replacedPXml)
      }
    }

    zip.setText('word/document.xml', docXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    await zip.save(outputPath)
    return { success: true, matchesCount, outputPath }
  }

  /**
   * Edit specific paragraph by index (update text, style, or remove).
   * @param {string} filePath
   * @param {number} paragraphIndex
   * @param {string} newText
   * @param {object} [options]
   * @returns {Promise<{success: boolean, outputPath: string}>}
   */
  async editParagraph(filePath, paragraphIndex, newText, options = {}) {
    const { outputPath = filePath, style = null, deleteParagraph = false } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    const pElements = extractElements(docXml, 'w:p')
    if (paragraphIndex < 0 || paragraphIndex >= pElements.length) {
      throw new Error(`Paragraph index out of range: ${paragraphIndex} (total: ${pElements.length})`)
    }

    const targetP = pElements[paragraphIndex]

    if (deleteParagraph) {
      docXml = docXml.replace(targetP.outerXml, '')
    } else {
      let pPrXml = ''
      const existingPPr = extractElements(targetP.outerXml, 'w:pPr')[0]
      if (style) {
        pPrXml = `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`
      } else if (existingPPr) {
        pPrXml = existingPPr.outerXml
      }

      const newPXml = `<w:p>${pPrXml}<w:r><w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r></w:p>`
      docXml = docXml.replace(targetP.outerXml, newPXml)
    }

    zip.setText('word/document.xml', docXml)
    await zip.save(outputPath)
    return { success: true, outputPath }
  }

  /**
   * Insert paragraph, heading, or break before/after specified index.
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<{success: boolean, outputPath: string}>}
   */
  async insert(filePath, options = {}) {
    const {
      outputPath = filePath,
      position = 'end',
      targetIndex = 0,
      text = '',
      style = 'Normal',
      isHeading = false,
      headingLevel = 1,
      pageBreak = false
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    let itemXml = ''
    if (pageBreak) {
      itemXml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    } else {
      const styleName = isHeading ? `Heading${headingLevel}` : style
      const pStyle = styleName !== 'Normal' ? `<w:pStyle w:val="${styleName}"/>` : ''
      itemXml = `<w:p><w:pPr>${pStyle}</w:pPr><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
    }

    if (position === 'start') {
      docXml = docXml.replace('<w:body>', `<w:body>\n${itemXml}`)
    } else if (position === 'end') {
      if (docXml.includes('<w:sectPr')) {
        docXml = docXml.replace('<w:sectPr', `${itemXml}\n<w:sectPr`)
      } else {
        docXml = docXml.replace('</w:body>', `${itemXml}\n</w:body>`)
      }
    } else {
      const pElements = extractElements(docXml, 'w:p')
      if (targetIndex < 0 || targetIndex >= pElements.length) {
        throw new Error(`Target paragraph index ${targetIndex} out of range`)
      }
      const targetP = pElements[targetIndex]
      if (position === 'before') {
        docXml = docXml.replace(targetP.outerXml, `${itemXml}\n${targetP.outerXml}`)
      } else if (position === 'after') {
        docXml = docXml.replace(targetP.outerXml, `${targetP.outerXml}\n${itemXml}`)
      }
    }

    zip.setText('word/document.xml', docXml)
    await zip.save(outputPath)
    return { success: true, outputPath }
  }

  /**
   * Insert, inspect, or modify tables in DOCX.
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
   */
  async table(filePath, options = {}) {
    const {
      outputPath = filePath,
      action = 'create',
      tableIndex = 0,
      rows = [],
      cell = null,
      position = 'end'
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    if (action === 'create') {
      const tblXml = this._buildTableXml(rows)
      if (position === 'end') {
        if (docXml.includes('<w:sectPr')) {
          docXml = docXml.replace('<w:sectPr', `${tblXml}\n<w:sectPr`)
        } else {
          docXml = docXml.replace('</w:body>', `${tblXml}\n</w:body>`)
        }
      } else {
        docXml = docXml.replace('<w:body>', `<w:body>\n${tblXml}`)
      }
      zip.setText('word/document.xml', docXml)
      await zip.save(outputPath)
      return { success: true, action: 'create', outputPath }
    }

    if (action === 'inspect') {
      const tblElements = extractElements(docXml, 'w:tbl')
      if (tableIndex < 0 || tableIndex >= tblElements.length) {
        throw new Error(`Table index ${tableIndex} not found`)
      }
      const tbl = tblElements[tableIndex]
      const rowElements = extractElements(tbl.outerXml, 'w:tr')
      const tableData = []

      for (const r of rowElements) {
        const cellElements = extractElements(r.outerXml, 'w:tc')
        const rowData = cellElements.map(c => extractTextFromXml(c.outerXml).trim())
        tableData.push(rowData)
      }

      return {
        tableIndex,
        rowCount: rowElements.length,
        colCount: tableData[0]?.length || 0,
        data: tableData
      }
    }

    if (action === 'addRow') {
      const tblElements = extractElements(docXml, 'w:tbl')
      if (tableIndex < 0 || tableIndex >= tblElements.length) {
        throw new Error(`Table index ${tableIndex} not found`)
      }
      const tbl = tblElements[tableIndex]
      const rowCells = rows[0] || []
      const newTrXml = this._buildRowXml(rowCells)

      const updatedTblXml = tbl.outerXml.replace('</w:tbl>', `${newTrXml}</w:tbl>`)
      docXml = docXml.replace(tbl.outerXml, updatedTblXml)

      zip.setText('word/document.xml', docXml)
      await zip.save(outputPath)
      return { success: true, action: 'addRow', outputPath }
    }

    if (action === 'setCell') {
      if (!cell) throw new Error('Missing cell specification: { row, col, value }')
      const tblElements = extractElements(docXml, 'w:tbl')
      if (tableIndex < 0 || tableIndex >= tblElements.length) {
        throw new Error(`Table index ${tableIndex} not found`)
      }
      const tbl = tblElements[tableIndex]
      const rowElements = extractElements(tbl.outerXml, 'w:tr')
      if (cell.row < 0 || cell.row >= rowElements.length) {
        throw new Error(`Row index ${cell.row} out of range`)
      }
      const targetTr = rowElements[cell.row]
      const cellElements = extractElements(targetTr.outerXml, 'w:tc')
      if (cell.col < 0 || cell.col >= cellElements.length) {
        throw new Error(`Column index ${cell.col} out of range`)
      }
      const targetTc = cellElements[cell.col]

      const newTcXml = `<w:tc><w:p><w:r><w:t>${escapeXml(cell.value)}</w:t></w:r></w:p></w:tc>`
      const updatedTrXml = targetTr.outerXml.replace(targetTc.outerXml, newTcXml)
      const updatedTblXml = tbl.outerXml.replace(targetTr.outerXml, updatedTrXml)
      docXml = docXml.replace(tbl.outerXml, updatedTblXml)

      zip.setText('word/document.xml', docXml)
      await zip.save(outputPath)
      return { success: true, action: 'setCell', outputPath }
    }

    throw new Error(`Unsupported table action: ${action}`)
  }

  /**
   * Validate DOCX integrity and XML structure.
   * @param {string} filePath
   * @returns {Promise<{valid: boolean, errors: string[], details: object}>}
   */
  async validate(filePath) {
    const errors = []
    let zip
    try {
      zip = await ZipArchive.fromFile(filePath)
    } catch (err) {
      return { valid: false, errors: [`ZIP corruption error: ${err.message}`], details: {} }
    }

    if (!zip.has('[Content_Types].xml')) {
      errors.push('Missing [Content_Types].xml')
    }
    if (!zip.has('word/document.xml')) {
      errors.push('Missing word/document.xml')
    }

    const docXml = zip.getText('word/document.xml')
    if (docXml) {
      if (!docXml.includes('<w:document') || !docXml.includes('</w:document>')) {
        errors.push('Corrupted root element in word/document.xml')
      }
      if (!docXml.includes('<w:body>') || !docXml.includes('</w:body>')) {
        errors.push('Corrupted body element in word/document.xml')
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      details: {
        fileSize: fs.statSync(filePath).size,
        fileEntries: zip.list().length
      }
    }
  }

  _buildTableXml(rows) {
    const trs = rows.map(r => this._buildRowXml(r)).join('')
    return `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>${trs}</w:tbl>`
  }

  _buildRowXml(cells) {
    const tcs = cells.map(c => {
      const val = typeof c === 'object' && c !== null ? c.value : c
      return `<w:tc><w:p><w:r><w:t>${escapeXml(val ?? '')}</w:t></w:r></w:p></w:tc>`
    }).join('')
    return `<w:tr>${tcs}</w:tr>`
  }

  _createFallbackBlankDocx() {
    const zip = new ZipArchive()
    zip.setText('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)

    zip.setText('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)

    zip.setText('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sectPr/>
  </w:body>
</w:document>`)

    return zip
  }
}
