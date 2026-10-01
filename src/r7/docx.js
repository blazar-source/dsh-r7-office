import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'
import {
  DOCX_STYLE_VOCABULARY,
  buildBorderEdge,
  buildParagraphProperties,
  buildParagraphXml,
  buildRunProperties,
  ensureStyleDefinitions,
  extractParagraphFormatting,
  extractTableFormatting,
  findStyleIdByName,
  parseNumberingDefinitions,
  parseStyleDefinitions,
  resolveStyleId
} from './docx-styles.js'
import { ensureListNumbering, listLists } from './docx-numbering.js'
import {
  appendToBodyEnd,
  applySectionSettings,
  collectSections,
  describeSection,
  insertSectionBreak as insertSectionBreakXml,
  listPageBreaks,
  replaceSection
} from './docx-sections.js'
import { insertImage as insertImageIntoPackage, listImages, probeImage } from './docx-media.js'
import {
  buildHyperlinkXml,
  ensureExternalRelationship,
  insertHyperlink as insertHyperlinkElement,
  listHyperlinks,
  removeHyperlink,
  setHyperlinkText
} from './docx-hyperlink.js'
import {
  buildPartXml,
  ensurePageNumberField,
  ensurePart,
  listParts,
  partText,
  removePart,
  setPartText
} from './docx-header-footer.js'
import {
  addColumn as addColumnOp,
  addRow as addRowOp,
  buildCellXml,
  describeTable,
  formatCell as formatCellOp,
  mergeCells,
  removeColumn as removeColumnOp,
  removeRow as removeRowOp,
  setColumnWidths,
  tableText,
  unmergeCells
} from './docx-table.js'
import {
  DOCUMENT_RELS_PART,
  ensureDefaultContentType,
  parseRelationships
} from './docx-parts.js'

const RELS_XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const EMPTY_RELS = `${RELS_XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`

/**
 * High-performance, style-preserving DOCX processing engine for R7-Office.
 *
 * The engine works directly on the OOXML package: `word/document.xml` is
 * edited as text at the exact boundary of the element that changes, and every
 * other member is left to the ZIP layer, which replays untouched members
 * byte-for-byte. That is what makes "read, change one paragraph, save a copy"
 * a safe operation on a document authored by R7-Office, Excel, Word or
 * LibreOffice alike.
 *
 * Feature modules keep the surface organized:
 *
 *   - `docx-styles.js`     — normalized formatting read/write vocabulary
 *   - `docx-sections.js`   — `<w:sectPr>`, page setup, page and section breaks
 *   - `docx-header-footer.js` — header/footer parts and field preservation
 *   - `docx-media.js`      — images: media part, relationship, drawing run
 *   - `docx-hyperlink.js`  — hyperlinks and their relationships
 *   - `docx-table.js`      — table geometry, merges and cell formatting
 *   - `docx-numbering.js`  — `word/numbering.xml` for bulleted/numbered lists
 *   - `docx-parts.js`      — relationships, content types, part naming
 */
export class DocxEngine {
  constructor(r7Adapter = null) {
    this.r7Adapter = r7Adapter || new R7Adapter()
  }

  /* ------------------------------------------------------------------ *
   * Reading
   * ------------------------------------------------------------------ */

  /**
   * Inspect document structure, outline, headings, paragraph count, and tables.
   *
   * Every extra section is opt-in so the default result is exactly what it was
   * before the formatting read-back existed.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {boolean} [options.includeStyles=false] - normalized formatting per outline paragraph
   * @param {boolean} [options.includeSections=false] - page setup and section breaks
   * @param {boolean} [options.includeHeadersFooters=false] - header/footer parts and their fields
   * @param {boolean} [options.includeHyperlinks=false]
   * @param {boolean} [options.includeImages=false]
   * @param {boolean} [options.includeLists=false]
   * @returns {Promise<object>}
   */
  async inspect(filePath, options = {}) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) {
      throw new Error('Invalid DOCX: word/document.xml not found')
    }

    const context = this._formattingContext(zip)
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
      if (options.includeStyles) {
        info.formatting = extractParagraphFormatting(p.outerXml, context)
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

    const result = {
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

    if (options.includeSections) result.sections = this._describeSections(zip, docXml)
    if (options.includeHeadersFooters) result.headersFooters = listParts(zip, docXml)
    if (options.includeHyperlinks) result.hyperlinks = this._hyperlinks(zip, docXml)
    if (options.includeImages) result.images = listImages(docXml, this._relationships(zip).byId)
    if (options.includeLists) result.lists = listLists(docXml, zip.getText('word/numbering.xml'))

    return result
  }

  /**
   * Read structured text, paragraphs, or tables from DOCX.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {'structured'|'markdown'|'text'} [options.format='structured']
   * @param {number} [options.fromParagraph=0]
   * @param {number} [options.count=100]
   * @param {string|null} [options.query=null]
   * @param {boolean} [options.includeStyles=false] - attach normalized `formatting` to each paragraph
   * @returns {Promise<object>}
   */
  async read(filePath, options = {}) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) {
      throw new Error('Invalid DOCX: word/document.xml not found')
    }

    const { format = 'structured', fromParagraph = 0, count = 100, query = null, includeStyles = false } = options
    const context = includeStyles ? this._formattingContext(zip) : null

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

      const entry = {
        index: i,
        style,
        text
      }
      if (includeStyles) entry.formatting = extractParagraphFormatting(p.outerXml, context)
      paragraphs.push(entry)
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
   * Normalized formatting of a document — the read-back an agent needs before
   * changing anything: paragraph styles, fonts, colours, alignment, indents,
   * spacing, lists and table cell styles.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {number} [options.fromParagraph=0]
   * @param {number} [options.count=200]
   * @param {boolean} [options.includeRuns=true]
   * @param {boolean} [options.includeTables=true]
   * @param {boolean} [options.includeSections=true]
   * @param {boolean} [options.includeHeadersFooters=true]
   * @param {boolean} [options.includeStylesList=false]
   * @returns {Promise<object>}
   */
  async formatting(filePath, options = {}) {
    const {
      fromParagraph = 0,
      count = 200,
      includeRuns = true,
      includeTables = true,
      includeSections = true,
      includeHeadersFooters = true,
      includeStylesList = false
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    const context = this._formattingContext(zip)
    const pElements = extractElements(docXml, 'w:p')
    const paragraphs = []

    for (let i = 0; i < pElements.length; i++) {
      const formatting = extractParagraphFormatting(pElements[i].outerXml, context)
      formatting.text = extractTextFromXml(pElements[i].outerXml)
      if (!includeRuns) delete formatting.runs
      paragraphs.push({ index: i, ...formatting })
    }

    const result = {
      filePath,
      paragraphCount: paragraphs.length,
      paragraphs: paragraphs.slice(fromParagraph, fromParagraph + count),
      returned: Math.max(0, Math.min(count, paragraphs.length - fromParagraph)),
      tableCount: extractElements(docXml, 'w:tbl').length,
      styleVocabulary: DOCX_STYLE_VOCABULARY
    }

    if (includeTables) {
      result.tables = extractElements(docXml, 'w:tbl').map((table, index) => ({
        index,
        ...extractTableFormatting(table.outerXml, context)
      }))
    }
    if (includeSections) result.sections = this._describeSections(zip, docXml)
    if (includeHeadersFooters) result.headersFooters = listParts(zip, docXml)
    if (includeStylesList) {
      result.styles = [...context.styles.values()].map(style => ({
        styleId: style.styleId,
        name: style.name,
        type: style.type,
        basedOn: style.basedOn,
        headingLevel: style.headingLevel
      }))
    }

    return result
  }

  /**
   * Plain-text view of the document's tables.
   * @param {string} filePath
   * @returns {Promise<Array<{index: number, rows: number, cols: number, data: string[][]}>>}
   */
  async tables(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')
    return extractElements(docXml, 'w:tbl').map((table, index) => ({
      index,
      rows: extractElements(table.outerXml, 'w:tr').length,
      cols: extractElements(extractElements(table.outerXml, 'w:tr')[0]?.outerXml ?? '', 'w:tc').length,
      data: tableText(table.outerXml)
    }))
  }

  /**
   * Every bulleted/numbered list paragraph, with its marker.
   * @param {string} filePath
   * @returns {Promise<Array<object>>}
   */
  async lists(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')
    return listLists(docXml, zip.getText('word/numbering.xml'))
  }

  /* ------------------------------------------------------------------ *
   * Creating
   * ------------------------------------------------------------------ */

  /**
   * Create a new DOCX document.
   *
   * @param {string} outputPath
   * @param {object} [options]
   * @param {string} [options.title]
   * @param {Array<string|object>} [options.paragraphs]
   * @param {Array<object>} [options.tables] - `{ rows, widthsTwips, borders, alignment }`
   * @param {object} [options.page] - page setup for the final section (orientation, margins, widthCm…)
   * @param {object|string} [options.header] - `{ text, pageNumber, alignment }` or a plain string
   * @param {object|string} [options.footer]
   * @param {Array<object>} [options.images] - `{ path|data, paragraphIndex, widthCm, heightCm, alt }`
   * @param {Array<object>} [options.hyperlinks] - `{ text, url, anchor, paragraphIndex, newParagraph }`
   * @param {boolean} [options.ensureStyles=true] - define the heading/hyperlink styles used
   * @param {boolean} [options.overwrite=false]
   * @returns {Promise<{success: boolean, path: string, stylesAdded?: string[], header?: object, footer?: object}>}
   */
  async create(outputPath, options = {}) {
    const {
      title,
      paragraphs = [],
      tables = [],
      overwrite = false,
      ensureStyles = true
    } = options

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

    const existingStyles = parseStyleDefinitions(zip.getText('word/styles.xml'))
    const normalStyleId = findStyleIdByName(existingStyles, 'Normal')
    const numberingIds = this._needsNumbering(paragraphs) ? ensureListNumbering(zip) : null
    const requestedStyles = []

    const resolveAndTrack = (requested) => {
      const resolved = resolveStyleId(existingStyles, requested)
      if (!resolved || resolved === 'Normal' || resolved === normalStyleId) return null
      requestedStyles.push(resolved)
      return resolved
    }

    // Build document body elements
    const bodyParts = []
    const normalizeParagraph = (spec) => {
      if (typeof spec === 'string') return { text: spec }
      const detail = { ...spec }
      if (detail.isHeading) detail.style = `Heading${detail.headingLevel || 1}`
      if (detail.list === 'bullet' || detail.list === 'bullets') {
        detail.list = { numId: numberingIds?.bullet, level: detail.listLevel || 0 }
      } else if (detail.list === 'number' || detail.list === 'numbered') {
        detail.list = { numId: numberingIds?.decimal, level: detail.listLevel || 0 }
      } else if (typeof detail.list === 'string') {
        throw new Error(`Unknown list type "${detail.list}". Use "bullet" or "number".`)
      }
      return detail
    }

    if (title) {
      const styleId = resolveAndTrack('Heading1')
      bodyParts.push(buildParagraphXml({
        text: title,
        style: styleId,
        alignment: 'center',
        bold: true,
        size: 18
      }, { styleId: styleId || undefined }))
    }

    for (const spec of paragraphs) {
      const detail = normalizeParagraph(spec)
      const styleId = resolveAndTrack(detail.style || 'Normal')
      // `pageBreak` with text means "start this paragraph on a new page"; only a
      // break with no text at all becomes a paragraph of its own.
      if (detail.pageBreak && !detail.text) {
        bodyParts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        continue
      }
      if (detail.pageBreak) detail.pageBreakBefore = true
      bodyParts.push(buildParagraphXml({ ...detail, style: styleId }, { styleId: styleId || undefined }))
    }

    // Build tables if requested
    for (const tbl of tables) {
      bodyParts.push(this._buildTableXml(tbl.rows || [], tbl))
    }

    // Replace initial empty paragraphs in template
    if (bodyParts.length > 0) {
      const joined = bodyParts.join('\n')
      // Clean out initial empty <w:p> from template
      docXml = docXml.replace(/<w:body>\s*<w:p><w:r><\/w:r><w:r><\/w:r><\/w:p>/, '<w:body>')
      docXml = docXml.replace(/<w:body>\s*<w:p\/>/, '<w:body>')
      docXml = appendToBodyEnd(docXml, joined)
      zip.setText('word/document.xml', docXml)
    }

    if (ensureStyles) {
      const styleIds = [...new Set(requestedStyles)]
      if (options.hyperlinks && options.hyperlinks.length > 0) styleIds.push('Hyperlink')
      const stylesXml = zip.getText('word/styles.xml')
      if (stylesXml) {
        const ensured = ensureStyleDefinitions(stylesXml, styleIds)
        if (ensured.added.length > 0) zip.setText('word/styles.xml', ensured.xml)
      }
    }

    // Page setup / sections
    if (options.page && Object.keys(options.page).length > 0) {
      docXml = this._applyFinalSection(zip, docXml, options.page)
      zip.setText('word/document.xml', docXml)
    }

    // Headers and footers
    const headerResult = this._applyHeaderFooterOption(zip, docXml, options.header, 'header')
    docXml = headerResult.docXml
    const footerResult = this._applyHeaderFooterOption(zip, docXml, options.footer, 'footer')
    docXml = footerResult.docXml
    zip.setText('word/document.xml', docXml)

    // Images and hyperlinks
    for (const image of options.images || []) {
      const inserted = await this._insertImageInto(zip, docXml, image)
      docXml = inserted.docXml
      zip.setText('word/document.xml', docXml)
    }
    for (const link of options.hyperlinks || []) {
      const inserted = this._insertHyperlinkInto(zip, docXml, link)
      docXml = inserted.docXml
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

  /* ------------------------------------------------------------------ *
   * Text editing
   * ------------------------------------------------------------------ */

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
   *
   * The paragraph keeps its `<w:pPr>` (and therefore its section break, list
   * numbering and spacing) unless a new style is requested, and the new run
   * inherits the previous first run's character formatting.
   *
   * @param {string} filePath
   * @param {number} paragraphIndex
   * @param {string} newText
   * @param {object} [options]
   * @returns {Promise<{success: boolean, outputPath: string}>}
   */
  async editParagraph(filePath, paragraphIndex, newText, options = {}) {
    const {
      outputPath = filePath,
      style = null,
      deleteParagraph = false,
      preserveFormatting = true,
      formatting = null
    } = options
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
      const existingPPr = extractElements(targetP.outerXml, 'w:pPr')[0]
      let pPrXml = ''
      if (style) {
        const styles = parseStyleDefinitions(zip.getText('word/styles.xml'))
        const styleId = resolveStyleId(styles, style)
        pPrXml = existingPPr
          ? `<w:pPr>${buildParagraphProperties({}, { styleId })}${existingPPr.innerXml.replace(/<w:pStyle(?=[\s/>])[^>]*\/?>/i, '')}</w:pPr>`
          : `<w:pPr><w:pStyle w:val="${escapeXml(styleId)}"/></w:pPr>`
      } else if (existingPPr) {
        pPrXml = existingPPr.outerXml
      }

      let runProperties = ''
      if (preserveFormatting) {
        const firstRun = extractElements(targetP.outerXml, 'w:r')[0]
        const rPr = firstRun ? extractElements(firstRun.outerXml, 'w:rPr')[0] : null
        if (rPr) runProperties = rPr.outerXml
      }
      if (formatting) {
        runProperties = `<w:rPr>${buildRunProperties(formatting)}</w:rPr>`
      }

      const newPXml = `<w:p>${pPrXml}<w:r>${runProperties}<w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r></w:p>`
      docXml = docXml.replace(targetP.outerXml, newPXml)
    }

    zip.setText('word/document.xml', docXml)
    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)
    return { success: true, outputPath }
  }

  /**
   * Insert a paragraph, heading, list item, image or break before/after a given
   * index.
   *
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<{success: boolean, outputPath: string, paragraphIndex?: number}>}
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

    const styles = parseStyleDefinitions(zip.getText('word/styles.xml'))
    let itemXml = ''
    if (pageBreak && !text) {
      itemXml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    } else {
      const requestedStyle = isHeading ? `Heading${headingLevel}` : style
      const styleId = resolveStyleId(styles, requestedStyle)
      const ensured = ensureStyleDefinitions(zip.getText('word/styles.xml'), [styleId])
      if (ensured.added.length > 0) zip.setText('word/styles.xml', ensured.xml)

      const detail = { ...options, text, style: styleId }
      if (detail.list === 'bullet' || detail.list === 'bullets' || detail.list === 'number' || detail.list === 'numbered') {
        const numbering = ensureListNumbering(zip)
        detail.list = {
          numId: detail.list === 'bullet' || detail.list === 'bullets' ? numbering.bullet : numbering.decimal,
          level: detail.listLevel || 0
        }
      }
      if (pageBreak) detail.pageBreakBefore = true
      itemXml = buildParagraphXml(detail, { styleId })
    }

    let insertedIndex = null
    if (position === 'start') {
      docXml = docXml.replace('<w:body>', `<w:body>\n${itemXml}`)
      insertedIndex = 0
    } else if (position === 'end') {
      docXml = appendToBodyEnd(docXml, itemXml)
      insertedIndex = extractElements(docXml, 'w:p').length - 1
    } else {
      const pElements = extractElements(docXml, 'w:p')
      if (targetIndex < 0 || targetIndex >= pElements.length) {
        throw new Error(`Target paragraph index ${targetIndex} out of range`)
      }
      const targetP = pElements[targetIndex]
      if (position === 'before') {
        docXml = docXml.replace(targetP.outerXml, `${itemXml}\n${targetP.outerXml}`)
        insertedIndex = targetIndex
      } else if (position === 'after') {
        docXml = docXml.replace(targetP.outerXml, `${targetP.outerXml}\n${itemXml}`)
        insertedIndex = targetIndex + 1
      }
    }

    zip.setText('word/document.xml', docXml)
    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)
    return { success: true, outputPath, paragraphIndex: insertedIndex }
  }

  /* ------------------------------------------------------------------ *
   * Tables
   * ------------------------------------------------------------------ */

  /**
   * Create, inspect, or modify tables in DOCX.
   *
   * Actions: `create`, `inspect`, `addRow`, `removeRow`, `addColumn`,
   * `removeColumn`, `setCell`, `formatCell`, `merge`, `unmerge`,
   * `setColumnWidths`.
   *
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

    const context = this._formattingContext(zip)
    const tblElements = () => extractElements(docXml, 'w:tbl')

    const resolveTable = () => {
      const tables = tblElements()
      if (tableIndex < 0 || tableIndex >= tables.length) {
        throw new Error(`Table index ${tableIndex} not found (document has ${tables.length} tables)`)
      }
      return tables[tableIndex]
    }

    const commit = async (tblXml) => {
      const table = resolveTable()
      docXml = docXml.slice(0, table.index) + tblXml + docXml.slice(table.index + table.outerXml.length)
      zip.setText('word/document.xml', docXml)
      const dir = path.dirname(outputPath)
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
      await zip.save(outputPath)
    }

    if (action === 'create') {
      const tblXml = this._buildTableXml(rows, options)
      if (position === 'end') {
        docXml = appendToBodyEnd(docXml, tblXml)
      } else {
        docXml = docXml.replace('<w:body>', `<w:body>\n${tblXml}`)
      }
      zip.setText('word/document.xml', docXml)
      const dir = path.dirname(outputPath)
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
      await zip.save(outputPath)
      return { success: true, action: 'create', outputPath }
    }

    if (action === 'inspect') {
      const tbl = resolveTable()
      const rowElements = extractElements(tbl.outerXml, 'w:tr')
      const tableData = rowElements.map(r => extractElements(r.outerXml, 'w:tc')
        .map(c => extractTextFromXml(c.outerXml).trim()))

      const result = {
        tableIndex,
        rowCount: rowElements.length,
        colCount: tableData[0]?.length || 0,
        data: tableData
      }
      if (options.includeStyles) result.formatting = describeTable(tbl.outerXml, context)
      return result
    }

    if (action === 'addRow') {
      const tbl = resolveTable()
      const newRowValues = options.values || rows[0] || []
      const updated = addRowOp(tbl.outerXml, {
        values: newRowValues,
        index: options.index,
        widthsTwips: options.widthsTwips,
        shading: options.shading,
        verticalAlign: options.verticalAlign,
        alignment: options.alignment
      })
      await commit(updated)
      return { success: true, action: 'addRow', outputPath }
    }

    if (action === 'removeRow') {
      const tbl = resolveTable()
      const rowIndex = options.rowIndex ?? options.index ?? 0
      await commit(removeRowOp(tbl.outerXml, Number(rowIndex)))
      return { success: true, action: 'removeRow', outputPath }
    }

    if (action === 'addColumn') {
      const tbl = resolveTable()
      const updated = addColumnOp(tbl.outerXml, {
        index: options.columnIndex ?? options.index ?? 0,
        values: options.values,
        widthTwips: options.widthTwips,
        shading: options.shading,
        verticalAlign: options.verticalAlign,
        alignment: options.alignment
      })
      await commit(updated)
      return { success: true, action: 'addColumn', outputPath }
    }

    if (action === 'removeColumn') {
      const tbl = resolveTable()
      await commit(removeColumnOp(tbl.outerXml, Number(options.columnIndex ?? options.index ?? 0)))
      return { success: true, action: 'removeColumn', outputPath }
    }

    if (action === 'merge') {
      const tbl = resolveTable()
      const spec = options.merge || cell || {}
      await commit(mergeCells(tbl.outerXml, {
        row: spec.row ?? 0,
        col: spec.col ?? 0,
        rows: spec.rows ?? spec.rowSpan ?? 1,
        cols: spec.cols ?? spec.colSpan ?? 1
      }))
      return { success: true, action: 'merge', outputPath }
    }

    if (action === 'unmerge') {
      const tbl = resolveTable()
      const spec = options.merge || cell || {}
      await commit(unmergeCells(tbl.outerXml, {
        row: spec.row ?? 0,
        col: spec.col ?? 0,
        axis: spec.axis || options.axis || 'all'
      }))
      return { success: true, action: 'unmerge', outputPath }
    }

    if (action === 'setColumnWidths') {
      const tbl = resolveTable()
      await commit(setColumnWidths(tbl.outerXml, {
        widthsTwips: options.widthsTwips,
        mode: options.mode,
        fixedLayout: options.fixedLayout
      }))
      return { success: true, action: 'setColumnWidths', outputPath }
    }

    if (action === 'setCell' || action === 'formatCell') {
      if (!cell) throw new Error('Missing cell specification: { row, col, value }')
      const tbl = resolveTable()
      const spec = {
        row: cell.row,
        col: cell.col,
        text: cell.value !== undefined ? cell.value : (action === 'formatCell' ? undefined : ''),
        shading: cell.shading,
        borders: cell.borders,
        clearBorders: cell.clearBorders,
        verticalAlign: cell.verticalAlign ?? cell.vAlign,
        alignment: cell.alignment ?? cell.align,
        widthTwips: cell.widthTwips,
        widthCm: cell.widthCm
      }

      // Keep the legacy behaviour for a plain text set: only the text changes,
      // every cell property (shading, borders, widths) survives.
      const hasFormatting = spec.shading !== undefined || spec.borders !== undefined
        || spec.clearBorders === true || spec.verticalAlign !== undefined || spec.alignment !== undefined
        || spec.widthTwips !== undefined || spec.widthCm !== undefined
      if (!hasFormatting) {
        const rowElements = extractElements(tbl.outerXml, 'w:tr')
        if (cell.row < 0 || cell.row >= rowElements.length) throw new Error(`Row index ${cell.row} out of range`)
        const cellElements = extractElements(rowElements[cell.row].outerXml, 'w:tc')
        if (cell.col < 0 || cell.col >= cellElements.length) throw new Error(`Column index ${cell.col} out of range`)
        const updatedCell = setCellTextPreservingProperties(cellElements[cell.col].outerXml, spec.text)
        const updatedRow = rowElements[cell.row].outerXml.replace(cellElements[cell.col].outerXml, updatedCell)
        await commit(tbl.outerXml.replace(rowElements[cell.row].outerXml, updatedRow))
      } else {
        await commit(formatCellOp(tbl.outerXml, spec))
      }
      return { success: true, action, outputPath }
    }

    throw new Error(`Unsupported table action: ${action}`)
  }

  /* ------------------------------------------------------------------ *
   * Page setup and sections
   * ------------------------------------------------------------------ */

  /**
   * Every section of the document, with page size, orientation, margins,
   * header/footer references and breaks.
   * @param {string} filePath
   * @returns {Promise<{filePath: string, sectionCount: number, sections: Array<object>, pageBreaks: Array<object>}>}
   */
  async sections(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')
    const sections = this._describeSections(zip, docXml)
    return {
      filePath,
      sectionCount: sections.length,
      sections,
      pageBreaks: listPageBreaks(docXml),
      vocabulary: {
        orientation: ['portrait', 'landscape'],
        sectionTypes: ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'],
        marginKeys: ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'],
        marginUnit: 'cm (numbers), or { twips } / { pt }',
        pageNumberFormats: ['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'decimalZero']
      }
    }
  }

  /**
   * Change page setup of one section, preserving everything not named.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {number} [options.sectionIndex=0] - 0-based index, or -1 for the final section
   * @param {string} [options.orientation] - portrait | landscape
   * @param {number} [options.widthCm] / [options.heightCm]
   * @param {number} [options.widthTwips] / [options.heightTwips]
   * @param {object} [options.margins] - cm per edge
   * @param {string} [options.type] - section break type
   * @param {number} [options.columns]
   * @param {number} [options.columnSpaceCm]
   * @param {boolean} [options.separator]
   * @param {boolean} [options.titlePg]
   * @param {number} [options.pageNumberStart]
   * @param {string} [options.pageNumberFormat]
   * @param {string} [options.outputPath]
   * @returns {Promise<{success: boolean, outputPath: string, sectionIndex: number, section: object}>}
   */
  async setSection(filePath, options = {}) {
    const { outputPath = filePath } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    const sections = collectSections(docXml)
    if (sections.length === 0) throw new Error('Invalid DOCX: the document has no <w:sectPr>')
    let index = options.sectionIndex ?? 0
    if (index === -1) index = sections.length - 1
    if (index < 0 || index >= sections.length) {
      throw new Error(`Section index out of range: ${index} (document has ${sections.length})`)
    }

    const updated = applySectionSettings(sections[index].outerXml, options)
    docXml = replaceSection(docXml, index, updated)
    zip.setText('word/document.xml', docXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    const after = this._describeSections(zip, docXml)
    return { success: true, outputPath, sectionIndex: index, section: after[index] }
  }

  /**
   * Insert a section break after a paragraph and apply page-setup changes to
   * the following section.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {number} options.afterParagraphIndex - last paragraph of the section being closed
   * @param {string} [options.type='nextPage']
   * @param {object} [options.page] - page setup for the new (following) section
   * @param {string} [options.outputPath]
   * @returns {Promise<{success: boolean, outputPath: string, sectionCount: number, breakParagraphIndex: number}>}
   */
  async insertSectionBreak(filePath, options = {}) {
    const { outputPath = filePath } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    const result = insertSectionBreakXml(docXml, options)
    docXml = result.xml
    zip.setText('word/document.xml', docXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    return {
      success: true,
      outputPath,
      sectionCount: collectSections(docXml).length,
      breakParagraphIndex: result.breakParagraphIndex
    }
  }

  /**
   * Page and section break positions.
   * @param {string} filePath
   * @returns {Promise<Array<object>>}
   */
  async pageBreaks(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')
    return listPageBreaks(docXml)
  }

  /* ------------------------------------------------------------------ *
   * Headers and footers
   * ------------------------------------------------------------------ */

  /**
   * Read or change headers and footers.
   *
   * Actions: `list`, `read`, `setText`, `create`, `addPageNumber`, `remove`.
   * Every text change keeps page-number and other fields intact.
   *
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
   */
  async headerFooter(filePath, options = {}) {
    const {
      outputPath = filePath,
      action = 'list',
      kind = 'header',
      type = 'default',
      sectionIndex = 0
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    const list = () => listParts(zip, docXml)

    if (action === 'list') {
      return { filePath, parts: list() }
    }

    if (action === 'read') {
      const parts = list()
      const part = parts.find(entry => entry.kind === kind && entry.type === type
        && entry.sectionIndexes.includes(Number(sectionIndex)))
      if (!part) {
        throw new Error(`No ${kind} of type "${type}" is referenced by section ${sectionIndex}. Use action "list" first.`)
      }
      const partXml = part.partName ? zip.getText(part.partName) : null
      return {
        filePath,
        kind,
        type,
        sectionIndex: Number(sectionIndex),
        partName: part.partName,
        exists: part.exists,
        text: part.text,
        fields: part.fields,
        hasPageNumberField: part.hasPageNumberField,
        paragraphCount: part.paragraphCount,
        paragraphs: partXml
          ? extractElements(partXml, 'w:p').map((paragraph, index) => ({
              index,
              text: partText(paragraph.outerXml)
            }))
          : []
      }
    }

    if (action === 'remove') {
      const result = removePart(zip, docXml, { kind, type, sectionIndex: Number(sectionIndex) })
      docXml = result.docXml
      zip.setText('word/document.xml', docXml)
      await zip.save(outputPath)
      return {
        success: true,
        action,
        outputPath,
        removedPart: result.removedPart,
        removedRelId: result.removedRelId
      }
    }

    if (action === 'create' || action === 'setText' || action === 'addPageNumber') {
      const field = options.pageNumber === true
        ? 'PAGE'
        : (typeof options.pageNumber === 'string' ? options.pageNumber : null)
      const ensured = ensurePart(zip, docXml, {
        kind,
        type,
        sectionIndex: Number(sectionIndex),
        alignment: options.alignment
      })
      docXml = ensured.docXml

      // The text and the field are applied to the part itself (whether it was
      // just created or already existed), so every action behaves identically.
      let partXml = zip.getText(ensured.partName) || buildPartXml(kind)
      if (options.text !== undefined && options.text !== null) {
        partXml = setPartText(partXml, options.text, { paragraphIndex: options.paragraphIndex ?? 0 }).xml
      }
      if (field) {
        const withField = ensurePageNumberField(partXml, {
          instruction: field,
          paragraphIndex: options.paragraphIndex ?? 0,
          position: options.position || 'end',
          separator: options.separator
        })
        partXml = withField.xml
      }
      zip.setText(ensured.partName, partXml)
      zip.setText('word/document.xml', docXml)
      await zip.save(outputPath)

      const entry = listParts(zip, docXml).find(part => part.partName === ensured.partName)
      return {
        success: true,
        action,
        outputPath,
        kind,
        type,
        sectionIndex: Number(sectionIndex),
        partName: ensured.partName,
        relId: ensured.relId,
        created: ensured.created,
        text: entry?.text ?? partText(partXml),
        fields: entry?.fields ?? [],
        hasPageNumberField: entry?.hasPageNumberField ?? false
      }
    }

    throw new Error(`Unsupported header/footer action: ${action}`)
  }

  /* ------------------------------------------------------------------ *
   * Images
   * ------------------------------------------------------------------ */

  /**
   * Insert a PNG/JPEG image.
   *
   * The display size keeps the image's own aspect ratio when only one
   * dimension is given; naming both uses exactly those values.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {string} [options.imagePath] - path to the image file
   * @param {string} [options.data] - base64 (optionally a data: URL) alternative
   * @param {string|number} [options.paragraphIndex='end'] - `end`, `start`, `new` or an index
   * @param {number} [options.widthCm] / [options.heightCm]
   * @param {number} [options.widthPx] / [options.heightPx]
   * @param {string} [options.alt]
   * @param {string} [options.name]
   * @param {string} [options.outputPath]
   * @returns {Promise<object>}
   */
  async insertImage(filePath, options = {}) {
    const { outputPath = filePath } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    this._ensureDocumentRels(zip)
    const inserted = await this._insertImageInto(zip, docXml, options)
    docXml = inserted.docXml
    zip.setText('word/document.xml', docXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    return {
      success: true,
      outputPath,
      mediaPartName: inserted.mediaPartName,
      relId: inserted.relId,
      format: inserted.format,
      pixelWidth: inserted.pixelWidth,
      pixelHeight: inserted.pixelHeight,
      widthCm: inserted.widthCm,
      heightCm: inserted.heightCm,
      widthEmu: inserted.cx,
      heightEmu: inserted.cy,
      aspectRatio: inserted.aspectRatio,
      scaled: inserted.scaled,
      paragraphIndex: inserted.paragraphIndex
    }
  }

  /**
   * Every image in the document.
   * @param {string} filePath
   * @returns {Promise<Array<object>>}
   */
  async images(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')
    return listImages(docXml, this._relationships(zip).byId)
  }

  /* ------------------------------------------------------------------ *
   * Hyperlinks
   * ------------------------------------------------------------------ */

  /**
   * Read, insert, retext or remove hyperlinks.
   *
   * Actions: `list`, `insert`, `setText`, `remove`.
   *
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
   */
  async hyperlink(filePath, options = {}) {
    const { outputPath = filePath, action = 'list' } = options
    const zip = await ZipArchive.fromFile(filePath)
    let docXml = zip.getText('word/document.xml')
    if (!docXml) throw new Error('Invalid DOCX: word/document.xml not found')

    if (action === 'list') {
      return { filePath, hyperlinks: this._hyperlinks(zip, docXml) }
    }

    if (action === 'insert') {
      if (!options.url && !options.anchor) throw new Error('A hyperlink needs a url or an anchor')
      if (!options.text) throw new Error('A hyperlink needs text')

      this._ensureDocumentRels(zip)
      let relId = null
      if (options.url) {
        const ensured = ensureExternalRelationship(zip.getText(DOCUMENT_RELS_PART), options.url)
        zip.setText(DOCUMENT_RELS_PART, ensured.xml)
        relId = ensured.id
      }

      const stylesXml = zip.getText('word/styles.xml')
      if (stylesXml) {
        const ensuredStyle = ensureStyleDefinitions(stylesXml, [options.style || 'Hyperlink'])
        if (ensuredStyle.added.length > 0) zip.setText('word/styles.xml', ensuredStyle.xml)
      }

      const hyperlinkXml = buildHyperlinkXml({
        text: options.text,
        relId,
        anchor: options.anchor,
        tooltip: options.tooltip,
        style: options.style,
        bold: options.bold,
        italic: options.italic,
        size: options.size,
        color: options.color
      })
      const inserted = insertHyperlinkElement(docXml, { ...options, hyperlinkXml })
      docXml = inserted.docXml
      zip.setText('word/document.xml', docXml)

      const dir = path.dirname(outputPath)
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
      await zip.save(outputPath)

      const links = this._hyperlinks(zip, docXml)
      return {
        success: true,
        action,
        outputPath,
        relId,
        paragraphIndex: inserted.paragraphIndex,
        hyperlinks: links
      }
    }

    if (action === 'setText') {
      const result = setHyperlinkText(docXml, options, this._relationships(zip).byId)
      zip.setText('word/document.xml', result.docXml)
      await zip.save(outputPath)
      return {
        success: true,
        action,
        outputPath,
        index: result.index,
        paragraphIndex: result.paragraphIndex,
        url: result.url
      }
    }

    if (action === 'remove') {
      const result = removeHyperlink(docXml, options, this._relationships(zip).byId)
      zip.setText('word/document.xml', result.docXml)
      await zip.save(outputPath)
      return { success: true, action, outputPath, index: result.index }
    }

    throw new Error(`Unsupported hyperlink action: ${action}`)
  }

  /* ------------------------------------------------------------------ *
   * Validation
   * ------------------------------------------------------------------ */

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

    const details = {
      fileSize: fs.statSync(filePath).size,
      fileEntries: zip.list().length
    }
    if (errors.length === 0) {
      details.parts = this._referencedParts(zip, docXml)
    }

    return {
      valid: errors.length === 0,
      errors,
      details
    }
  }

  /* ------------------------------------------------------------------ *
   * Internals
   * ------------------------------------------------------------------ */

  /** Build a table element, honouring optional widths and borders. */
  _buildTableXml(rows, options = {}) {
    const widths = options.widthsTwips
      || (Array.isArray(options.widths) ? options.widths : null)
    const trs = rows.map(r => this._buildRowXml(r, widths)).join('')

    const borders = options.borders === false
      ? ''
      : (options.borders && typeof options.borders === 'object'
        ? this._buildTableBorders(options.borders)
        : '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>')

    const alignment = options.alignment ? `<w:jc w:val="${options.alignment}"/>` : ''
    const grid = widths && widths.length > 0
      ? `<w:tblGrid>${widths.map(width => `<w:gridCol w:w="${Math.round(Number(width))}"/>`).join('')}</w:tblGrid>`
      : ''
    const width = widths && widths.length > 0
      ? `<w:tblW w:w="${widths.reduce((sum, value) => sum + Number(value), 0)}" w:type="dxa"/>`
      : ''

    return `<w:tbl><w:tblPr>${width}${alignment}${borders}<w:tblLayout w:type="fixed"/></w:tblPr>${grid}${trs}</w:tbl>`
  }

  /** Build a table borders element from normalized edge specs. */
  _buildTableBorders(borders) {
    const edges = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .filter(edge => borders[edge] !== undefined && borders[edge] !== null)
    if (edges.length === 0) return ''
    return `<w:tblBorders>${edges.map(edge => buildBorderEdge(edge, borders[edge])).join('')}</w:tblBorders>`
  }

  /** Build a row element; object cells carry their own formatting. */
  _buildRowXml(cells, widths = null) {
    const tcs = cells.map((c, index) => {
      if (typeof c === 'object' && c !== null) {
        const spec = { ...c }
        if (spec.value !== undefined && spec.text === undefined) spec.text = spec.value
        if (spec.widthTwips === undefined && widths && widths[index] !== undefined) spec.widthTwips = widths[index]
        return buildCellXml(spec)
      }
      return buildCellXml({
        text: c ?? '',
        widthTwips: widths && widths[index] !== undefined ? widths[index] : undefined
      })
    }).join('')
    return `<w:tr>${tcs}</w:tr>`
  }

  /** A blank package used when no R7 template is installed. */
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

    zip.setText(DOCUMENT_RELS_PART, EMPTY_RELS)

    zip.setText('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sectPr/>
  </w:body>
</w:document>`)

    return zip
  }

  /** Parse the document relationships part. */
  _relationships(zip) {
    return parseRelationships(zip.getText(DOCUMENT_RELS_PART) || '')
  }

  /** Create the document relationships part when a package lacks it. */
  _ensureDocumentRels(zip) {
    if (!zip.has(DOCUMENT_RELS_PART)) {
      zip.setText(DOCUMENT_RELS_PART, EMPTY_RELS)
      const contentTypes = zip.getText('[Content_Types].xml')
      if (contentTypes) {
        const ensured = ensureDefaultContentType(
          contentTypes,
          'rels',
          'application/vnd.openxmlformats-package.relationships+xml'
        )
        if (ensured.added) zip.setText('[Content_Types].xml', ensured.xml)
      }
    }
    return zip.getText(DOCUMENT_RELS_PART)
  }

  /** Formatting context (numbering + style definitions) for a package. */
  _formattingContext(zip) {
    return {
      numbering: parseNumberingDefinitions(zip.getText('word/numbering.xml')),
      styles: parseStyleDefinitions(zip.getText('word/styles.xml'))
    }
  }

  /** Normalized sections of a document. */
  _describeSections(zip, docXml) {
    const relationships = this._relationships(zip)
    return collectSections(docXml).map(section => ({
      index: section.index,
      kind: section.kind,
      paragraphIndex: section.paragraphIndex,
      ...describeSection(section.outerXml, { relationships: relationships.byId })
    }))
  }

  /** Hyperlinks, with their relationships resolved. */
  _hyperlinks(zip, docXml) {
    return listHyperlinks(docXml, this._relationships(zip).byId)
  }

  /** Parts referenced by the document, for a health report. */
  _referencedParts(zip, docXml) {
    const relationships = this._relationships(zip)
    const memberNames = new Set(zip.list())
    const parts = {}
    for (const rel of relationships.list) {
      const target = rel.target.startsWith('/')
        ? rel.target.replace(/^\/+/, '')
        : (rel.target.startsWith('word/') ? rel.target : `word/${rel.target}`)
      if (rel.targetMode === 'External') continue
      parts[rel.id] = { target, present: memberNames.has(target) }
    }
    const sections = collectSections(docXml)
    return {
      relationships: parts,
      sections: sections.length,
      headersFooters: listParts(zip, docXml).map(part => ({
        kind: part.kind,
        type: part.type,
        partName: part.partName,
        exists: part.exists
      }))
    }
  }

  /** Apply page-setup options to the final section. */
  _applyFinalSection(zip, docXml, page) {
    const sections = collectSections(docXml)
    if (sections.length === 0) return docXml
    const index = sections.length - 1
    const updated = applySectionSettings(sections[index].outerXml, page)
    return replaceSection(docXml, index, updated)
  }

  /** Apply the `header`/`footer` create option. */
  _applyHeaderFooterOption(zip, docXml, spec, kind) {
    if (spec === undefined || spec === null) return { docXml, partName: null }
    const detail = typeof spec === 'string' ? { text: spec } : spec
    const field = detail.pageNumber === true
      ? 'PAGE'
      : (typeof detail.pageNumber === 'string' ? detail.pageNumber : null)
    const ensured = ensurePart(zip, docXml, {
      kind,
      type: detail.type || 'default',
      sectionIndex: detail.sectionIndex ?? 0,
      text: detail.text,
      field,
      alignment: detail.alignment
    })
    return { docXml: ensured.docXml, partName: ensured.partName }
  }

  /** Insert an image described by an options object (path or base64 data). */
  async _insertImageInto(zip, docXml, options) {
    const buffer = this._imageBuffer(options)
    const probed = probeImage(buffer)
    const result = insertImageIntoPackage(zip, docXml, {
      buffer,
      fileName: options.imagePath || options.path || options.fileName,
      paragraphIndex: options.paragraphIndex ?? 'end',
      widthCm: options.widthCm,
      heightCm: options.heightCm,
      widthPx: options.widthPx,
      heightPx: options.heightPx,
      widthPt: options.widthPt,
      heightPt: options.heightPt,
      alt: options.alt,
      name: options.name,
      extension: options.extension,
      contentType: options.contentType
    })
    return { ...result, format: result.format || probed.format }
  }

  /** Decode the image payload from `imagePath`/`path` or base64 `data`. */
  _imageBuffer(options) {
    if (options.buffer && Buffer.isBuffer(options.buffer)) return options.buffer
    const filePath = options.imagePath || options.path
    if (filePath) {
      if (!fs.existsSync(filePath)) throw new Error(`Image file does not exist: ${filePath}`)
      return fs.readFileSync(filePath)
    }
    if (options.data) {
      const base64 = String(options.data).replace(/^data:[^;]+;base64,/, '')
      return Buffer.from(base64, 'base64')
    }
    throw new Error('insertImage requires imagePath, data (base64) or buffer')
  }

  /** Insert a hyperlink described by an options object. */
  _insertHyperlinkInto(zip, docXml, options) {
    this._ensureDocumentRels(zip)
    let relId = null
    if (options.url) {
      const ensured = ensureExternalRelationship(zip.getText(DOCUMENT_RELS_PART), options.url)
      zip.setText(DOCUMENT_RELS_PART, ensured.xml)
      relId = ensured.id
    }

    const stylesXml = zip.getText('word/styles.xml')
    if (stylesXml) {
      const ensuredStyle = ensureStyleDefinitions(stylesXml, [options.style || 'Hyperlink'])
      if (ensuredStyle.added.length > 0) zip.setText('word/styles.xml', ensuredStyle.xml)
    }

    const hyperlinkXml = buildHyperlinkXml({
      text: options.text,
      relId,
      anchor: options.anchor,
      tooltip: options.tooltip,
      style: options.style,
      bold: options.bold,
      italic: options.italic,
      size: options.size,
      color: options.color
    })
    const inserted = insertHyperlinkElement(docXml, { ...options, hyperlinkXml })
    return { ...inserted, relId, paragraphIndex: inserted.paragraphIndex }
  }

  /** Whether any paragraph spec asks for a list. */
  _needsNumbering(paragraphs) {
    return paragraphs.some(spec => spec && typeof spec === 'object' && (spec.list === 'bullet'
      || spec.list === 'bullets' || spec.list === 'number' || spec.list === 'numbered'))
  }
}

/** Local helpers that keep the engine's imports focused. */

/** Set a cell's text while keeping the cell's `<w:tcPr>` and run properties. */
function setCellTextPreservingProperties(cellXml, text) {
  const paragraphs = extractElements(cellXml, 'w:p')
  if (paragraphs.length === 0) {
    const tcPr = extractElements(cellXml, 'w:tcPr')[0]
    const at = tcPr ? tcPr.index + tcPr.outerXml.length : (/<w:tc\b[^>]*>/.exec(cellXml)?.[0].length ?? 0)
    return cellXml.slice(0, at) + `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>` + cellXml.slice(at)
  }
  const target = paragraphs[0]
  const pPr = extractElements(target.outerXml, 'w:pPr')[0]
  const firstRun = extractElements(target.outerXml, 'w:r')[0]
  const rPr = firstRun ? extractElements(firstRun.outerXml, 'w:rPr')[0] : null
  const paragraph = `<w:p>${pPr ? pPr.outerXml : ''}<w:r>${rPr ? rPr.outerXml : ''}<w:t>${escapeXml(text)}</w:t></w:r></w:p>`
  return cellXml.slice(0, target.index) + paragraph + cellXml.slice(target.index + target.outerXml.length)
}
