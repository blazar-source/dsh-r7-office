import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { escapeXml, unescapeXml, extractTextFromXml, extractElements, getAttribute } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'

/**
 * High-performance PPTX presentation engine for R7-Office.
 */
export class PptxEngine {
  constructor(r7Adapter = null) {
    this.r7Adapter = r7Adapter || new R7Adapter()
  }

  /**
   * Inspect PPTX presentation: slide count, titles, notes, shapes.
   * @param {string} filePath
   * @returns {Promise<object>}
   */
  async inspect(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const presXml = zip.getText('ppt/presentation.xml')
    if (!presXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const slides = []
    const sldIdElements = extractElements(presXml, 'p:sldId')

    for (let i = 0; i < sldIdElements.length; i++) {
      const slidePath = `ppt/slides/slide${i + 1}.xml`
      let title = `Slide ${i + 1}`
      let textItems = []

      if (zip.has(slidePath)) {
        const slideXml = zip.getText(slidePath)
        const tElements = extractElements(slideXml, 'a:t')
        textItems = tElements.map(t => extractTextFromXml(t.outerXml).trim()).filter(Boolean)
        if (textItems.length > 0) {
          title = textItems[0]
        }
      }

      slides.push({
        index: i,
        slideNumber: i + 1,
        title,
        textCount: textItems.length,
        preview: textItems.slice(0, 3).join(' | ')
      })
    }

    return {
      type: 'pptx',
      filePath,
      slidesCount: slides.length,
      slides
    }
  }

  /**
   * Read all text content from a PPTX slide or whole presentation.
   * @param {string} filePath
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async read(filePath, options = {}) {
    const { slideIndex = null } = options
    const zip = await ZipArchive.fromFile(filePath)
    const presXml = zip.getText('ppt/presentation.xml')
    if (!presXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const sldIdElements = extractElements(presXml, 'p:sldId')
    const slideResults = []

    for (let i = 0; i < sldIdElements.length; i++) {
      if (slideIndex !== null && slideIndex !== i) continue

      const slidePath = `ppt/slides/slide${i + 1}.xml`
      const texts = []

      if (zip.has(slidePath)) {
        const slideXml = zip.getText(slidePath)
        const pElements = extractElements(slideXml, 'a:p')
        for (const p of pElements) {
          const t = extractTextFromXml(p.outerXml).trim()
          if (t) texts.push(t)
        }
      }

      slideResults.push({
        slideNumber: i + 1,
        index: i,
        paragraphs: texts
      })
    }

    return {
      filePath,
      slides: slideResults
    }
  }

  /**
   * Create a new PPTX presentation.
   * @param {string} outputPath
   * @param {object} [options]
   * @returns {Promise<{success: boolean, path: string}>}
   */
  async create(outputPath, options = {}) {
    const { title = 'Новая презентация', overwrite = false } = options

    if (fs.existsSync(outputPath) && overwrite !== true) {
      throw new Error(
        `Refusing to overwrite: ${outputPath} already exists. Pass overwrite: true to replace it.`
      )
    }

    const tplPath = await this.r7Adapter.getTemplatePath('pptx')
    let zip

    if (tplPath && fs.existsSync(tplPath)) {
      zip = await ZipArchive.fromFile(tplPath)
    } else {
      zip = this._createFallbackBlankPptx()
    }

    if (title && zip.has('ppt/slides/slide1.xml')) {
      let slide1Xml = zip.getText('ppt/slides/slide1.xml')
      slide1Xml = this._setSlideTitle(slide1Xml, title)
      zip.setText('ppt/slides/slide1.xml', slide1Xml)
    }

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    await zip.save(outputPath)
    return { success: true, path: outputPath }
  }

  /**
   * Append a slide to an existing presentation.
   *
   * Every part the OOXML part graph requires is registered: the slide itself,
   * its relationship part (copying the base slide's layout link), the
   * `[Content_Types].xml` override, the `<p:sldId>` entry and the presentation
   * relationship. Slides that already exist are left untouched, so appending
   * can never destroy a deck.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {string} [options.title] - Title text for the new slide.
   * @param {number} [options.baseSlideIndex] - 0-based slide to clone (default 0).
   * @param {string} [options.outputPath]
   * @returns {Promise<{success: boolean, slideNumber: number, slideCount: number, outputPath: string}>}
   */
  async addSlide(filePath, options = {}) {
    const { outputPath = filePath, title = '', baseSlideIndex = 0 } = options

    const zip = await ZipArchive.fromFile(filePath)
    const presentationPath = 'ppt/presentation.xml'
    const presentationXml = zip.getText(presentationPath)
    if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    // Fresh part number: one past the highest existing slideN.xml.
    const partNumbers = zip.list()
      .map((entry) => {
        const m = entry.match(/^ppt\/slides\/slide(\d+)\.xml$/)
        return m ? Number(m[1]) : 0
      })
      .filter(Boolean)
    const newPart = partNumbers.length > 0 ? Math.max(...partNumbers) + 1 : 1

    const basePart = partNumbers.includes(baseSlideIndex + 1)
      ? baseSlideIndex + 1
      : (partNumbers.length > 0 ? Math.min(...partNumbers) : 1)
    const basePath = `ppt/slides/slide${basePart}.xml`

    let slideXml = zip.getText(basePath) || this._blankSlideXml()
    if (title) slideXml = this._setSlideTitle(slideXml, title)

    // The cloned slide keeps the base slide's layout link but must not inherit
    // its notes-slide relationship, which belongs to the base slide alone.
    const baseRelsPath = `ppt/slides/_rels/slide${basePart}.xml.rels`
    let slideRels = zip.getText(baseRelsPath)
    if (slideRels) {
      const layoutRel = extractElements(slideRels, 'Relationship')
        .find((r) => /\/slideLayout$/.test(getAttribute(r.outerXml, 'Type') || ''))
      slideRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + (layoutRel ? layoutRel.outerXml : '')
        + '</Relationships>'
    } else {
      slideRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    }

    // Presentation relationship for the new slide.
    const relsPath = 'ppt/_rels/presentation.xml.rels'
    let relsXml = zip.getText(relsPath)
    if (!relsXml) {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    }
    const relIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
    const newRelId = `rId${relIds.length > 0 ? Math.max(...relIds) + 1 : 1}`
    relsXml = relsXml.replace(
      '</Relationships>',
      `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newPart}.xml"/></Relationships>`
    )

    // Content type override for the new slide part.
    const contentTypesPath = '[Content_Types].xml'
    let contentTypes = zip.getText(contentTypesPath)
    if (!contentTypes) throw new Error('Invalid PPTX: [Content_Types].xml not found')
    if (!contentTypes.includes(`/ppt/slides/slide${newPart}.xml`)) {
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override PartName="/ppt/slides/slide${newPart}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
      )
    }

    // <p:sldId> ids must be unique within the presentation.
    const sldIds = [...presentationXml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => Number(m[1]))
    const newSldId = sldIds.length > 0 ? Math.max(...sldIds) + 1 : 256
    let updatedPresentation

    if (presentationXml.includes('</p:sldIdLst>')) {
      updatedPresentation = presentationXml.replace(
        '</p:sldIdLst>',
        `<p:sldId id="${newSldId}" r:id="${newRelId}"/></p:sldIdLst>`
      )
    } else if (presentationXml.includes('<p:sldIdLst/>')) {
      updatedPresentation = presentationXml.replace(
        '<p:sldIdLst/>',
        `<p:sldIdLst><p:sldId id="${newSldId}" r:id="${newRelId}"/></p:sldIdLst>`
      )
    } else {
      // A presentation with no slide list at all: insert one after the master
      // list, which the schema requires to come first.
      const sldIdLst = `<p:sldIdLst><p:sldId id="${newSldId}" r:id="${newRelId}"/></p:sldIdLst>`
      updatedPresentation = /<\/p:sldMasterIdLst>/.test(presentationXml)
        ? presentationXml.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst>${sldIdLst}`)
        : presentationXml.replace('<p:presentation', `<p:presentation`).replace(/>/, `>${sldIdLst}`)
    }

    zip.setBuffer(`ppt/slides/slide${newPart}.xml`, Buffer.from(slideXml, 'utf8'))
    zip.setText(`ppt/slides/_rels/slide${newPart}.xml.rels`, slideRels)
    zip.setText(relsPath, relsXml)
    zip.setText(contentTypesPath, contentTypes)
    zip.setText(presentationPath, updatedPresentation)
    this._syncSlideCount(zip, Math.max(...partNumbers, newPart))

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)

    const slideCountAfter = extractElements(updatedPresentation, 'p:sldId').length
    return {
      success: true,
      slideNumber: newPart,
      slideCount: slideCountAfter,
      outputPath
    }
  }

  /**
   * Keep the extended-properties slide count consistent when present. R7 and
   * PowerPoint tolerate a stale value, but a correct one keeps validators quiet.
   */
  _syncSlideCount(zip, count) {
    const appPath = 'docProps/app.xml'
    const appXml = zip.getText(appPath)
    if (!appXml || !/<Slides>\d+<\/Slides>/.test(appXml)) return
    zip.setText(appPath, appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${count}</Slides>`))
  }

  _blankSlideXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
      + '<p:cSld><p:spTree>'
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
      + '<p:grpSpPr/>'
      + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>'
      + '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU" dirty="0"/></a:p></p:txBody></p:sp>'
      + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  }

  /**
   * Edit text in slide shapes by matching search string or target title.
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
   */
  async editSlide(filePath, options = {}) {
    const {
      outputPath = filePath,
      slideIndex = 0,
      search = null,
      replace = null,
      title = null
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`

    if (!zip.has(slidePath)) {
      throw new Error(`Slide not found: ${slidePath}`)
    }

    let slideXml = zip.getText(slidePath)
    let matchesCount = 0

    if (title) {
      slideXml = this._setSlideTitle(slideXml, title)
      matchesCount++
    }

    if (search && replace !== null) {
      const tElements = extractElements(slideXml, 'a:t')
      for (const t of tElements) {
        const text = extractTextFromXml(t.outerXml)
        if (text.includes(search)) {
          const newText = text.replaceAll(search, replace)
          slideXml = slideXml.replace(t.outerXml, `<a:t>${escapeXml(newText)}</a:t>`)
          matchesCount++
        }
      }
    }

    zip.setText(slidePath, slideXml)

    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })

    await zip.save(outputPath)
    return { success: true, matchesCount, outputPath }
  }

  /**
   * Validate PPTX integrity.
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
    if (!zip.has('ppt/presentation.xml')) errors.push('Missing ppt/presentation.xml')
    if (!zip.has('ppt/slides/slide1.xml')) errors.push('Missing ppt/slides/slide1.xml')

    return {
      valid: errors.length === 0,
      errors,
      details: {
        fileSize: fs.statSync(filePath).size,
        entries: zip.list().length
      }
    }
  }

  _setSlideTitle(slideXml, title) {
    const tElements = extractElements(slideXml, 'a:t')
    if (tElements.length > 0) {
      return slideXml.replace(tElements[0].outerXml, `<a:t>${escapeXml(title)}</a:t>`)
    }

    // If template has empty <a:p> in title shape
    const spElements = extractElements(slideXml, 'p:sp')
    if (spElements.length > 0) {
      const firstSp = spElements[0]
      const pElements = extractElements(firstSp.outerXml, 'a:p')
      if (pElements.length > 0) {
        const firstP = pElements[0]
        const newPXml = `<a:p><a:r><a:rPr lang="ru-RU"/><a:t>${escapeXml(title)}</a:t></a:r></a:p>`
        const updatedSp = firstSp.outerXml.replace(firstP.outerXml, newPXml)
        return slideXml.replace(firstSp.outerXml, updatedSp)
      }
    }
    return slideXml
  }

  _createFallbackBlankPptx() {
    const zip = new ZipArchive()
    zip.setText('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)

    zip.setText('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`)

    zip.setText('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`)

    zip.setText('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`)

    zip.setText('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Новая презентация</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`)

    return zip
  }
}
