import fs from 'node:fs'
import path from 'node:path'
import { ZipArchive } from '../shared/zip.js'
import { extractElements, getAttribute } from '../shared/xml.js'
import { R7Adapter } from './adapter.js'
import {
  CONTENT_TYPE,
  REL,
  ensureOverride,
  firstElement,
  openingTag,
  readRelationships,
  relsPathFor,
  slideLayouts,
  slideParts,
  slideSize
} from './pptx-util.js'
import {
  buildPicture,
  buildShape,
  buildTextBody,
  ensureNormAutofit,
  normalizeAnchor,
  patchShapeProperties,
  patchTextBodyAttribute,
  patchTextBodyFont,
  patchTextBodyParagraphProperties,
  patchOneParagraph,
  resolvePreset,
  SHAPE_PRESETS
} from './pptx-build.js'
import { readSlide as readSlideModel, collectShapeElements } from './pptx-read.js'
import {
  addImageRelationship,
  dropImageRelationship,
  fitImage,
  listMedia,
  loadImage,
  patchPicture,
  replaceMedia,
  retargetImageRelationship,
  storeMedia
} from './pptx-image.js'

/**
 * High-performance PPTX presentation engine for R7-Office.
 *
 * Design rules this engine holds to:
 *
 *  - Slides are addressed by their *position in the deck*, resolved through
 *    `ppt/presentation.xml.rels`, never by guessing `slide<index + 1>.xml`.
 *    Deleting slide 2 of 4 leaves a part-numbering gap, and every read and
 *    write has to keep working after that.
 *  - An existing object is patched, not rebuilt. Restyling one text run must
 *    leave the gradient, the shadow and the hyperlink the author set alone.
 *  - Nothing outside the addressed part is touched: layouts, master, theme,
 *    notes, media and every other slide keep their original bytes.
 */
export class PptxEngine {
  constructor(r7Adapter = null) {
    this.r7Adapter = r7Adapter || new R7Adapter()
  }

  // =========================================================== reading

  /**
   * Inspect PPTX presentation: slide count, titles, notes, layouts.
   * @param {string} filePath
   * @returns {Promise<object>}
   */
  async inspect(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const descriptors = slideParts(zip)
    const layouts = slideLayouts(zip)
    const size = slideSize(zip)

    const slides = descriptors.map((descriptor, index) => {
      const slideXml = zip.getText(descriptor.partPath) || ''
      const texts = extractElements(slideXml, 'a:t')
        .map((t) => decodeText(t.outerXml))
        .filter(Boolean)
      const rels = readRelationships(zip, descriptor.partPath)
      const layoutRel = rels.find((r) => r.type === REL.slideLayout)
      const layout = layoutRel ? layouts.find((l) => l.partPath === layoutRel.partPath) : null
      const tree = extractElements(slideXml, 'p:spTree')[0]

      return {
        index,
        slideNumber: index + 1,
        partPath: descriptor.partPath,
        title: texts.length > 0 ? texts[0] : `Slide ${index + 1}`,
        textCount: texts.length,
        preview: texts.slice(0, 3).join(' | '),
        layout: layout ? { name: layout.name, type: layout.type, index: layout.index } : null,
        hasNotes: rels.some((r) => r.type === REL.notesSlide),
        objectCount: tree ? collectShapeElements(tree.innerXml).length : 0
      }
    })

    return {
      type: 'pptx',
      filePath,
      slidesCount: slides.length,
      slideSize: size,
      layouts: layouts.map((l) => ({ index: l.index, name: l.name, type: l.type, partPath: l.partPath })),
      mediaCount: listMedia(zip).length,
      slides
    }
  }

  /**
   * Read a slide as a normalized structure.
   *
   * This is the primary read entry point: it reports every object with its
   * geometry, text, font, fill, stroke, alignment and paragraphs, resolving the
   * geometry and typography a placeholder inherits from its layout.
   *
   * @param {string} filePath
   * @param {number} [slideIndex=0] - 0-based position in the deck.
   * @param {object} [options]
   * @param {boolean} [options.includeInherited=true] - also report layout
   *   placeholders that exist only in the layout.
   * @param {boolean} [options.includeRaw=false] - attach the raw OOXML of each
   *   object, for debugging.
   * @returns {Promise<object>}
   */
  async readSlide(filePath, slideIndex = 0, options = {}) {
    const zip = await ZipArchive.fromFile(filePath)
    const descriptors = slideParts(zip)
    if (slideIndex < 0 || slideIndex >= descriptors.length) {
      throw new Error(
        `Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s) (0..${descriptors.length - 1})`
      )
    }

    const layouts = slideLayouts(zip)
    const size = slideSize(zip)
    const all = descriptors.map((descriptor) => readSlideModel(zip, descriptor, {
      layouts,
      size,
      includeInherited: options.includeInherited !== false
    }))

    const selected = all[slideIndex]
    if (options.includeRaw !== true) {
      // The raw markup stays on every object on purpose: the write path locates
      // an object by searching for the exact XML it parsed, which is what makes
      // "edit this object" a substring operation rather than a second, possibly
      // divergent parse.
    }
    return {
      filePath,
      slideCount: descriptors.length,
      layouts: layouts.map((l) => ({ index: l.index, name: l.name, type: l.type, partPath: l.partPath })),
      slide: selected,
      // Reporting every slide's object list makes a two-call workflow (read a
      // slide, edit a slide) unnecessary for the common "what is on this deck"
      // question, at the cost of one extra pass over tiny XML parts.
      slides: all.map((slide, index) => (index === slideIndex
        ? null
        : { index: slide.index, slideNumber: slide.slideNumber, name: slide.name, layout: slide.layout, objectCount: slide.objects.length }))
        .filter(Boolean)
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
    const descriptors = slideParts(zip)

    const slides = []
    for (const descriptor of descriptors) {
      if (slideIndex !== null && slideIndex !== descriptor.index) continue
      const slideXml = zip.getText(descriptor.partPath)
      const paragraphs = []
      for (const p of extractElements(slideXml, 'a:p')) {
        const text = decodeText(p.outerXml).trim()
        if (text) paragraphs.push(text)
      }
      slides.push({
        slideNumber: descriptor.index + 1,
        index: descriptor.index,
        partPath: descriptor.partPath,
        paragraphs
      })
    }

    return {
      filePath,
      slides
    }
  }

  /**
   * List the slide layouts available in a deck.
   * @param {string} filePath
   * @returns {Promise<object>}
   */
  async listLayouts(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    const layouts = slideLayouts(zip)
    const size = slideSize(zip)
    return {
      filePath,
      slideSize: size,
      layouts: layouts.map((l) => ({
        index: l.index,
        name: l.name,
        type: l.type,
        partPath: l.partPath,
        placeholders: l.placeholders.map((p) => ({
          type: p.type, idx: p.idx, name: p.name, x: p.x, y: p.y, width: p.width, height: p.height
        }))
      }))
    }
  }

  /** List the media parts in a deck. */
  async listMedia(filePath) {
    const zip = await ZipArchive.fromFile(filePath)
    return { filePath, media: listMedia(zip) }
  }

  // ============================================================ creation

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

    return this._write(zip, outputPath, { success: true, path: outputPath })
  }

  /**
   * Append a slide to a presentation, based on an existing layout.
   *
   * Every part the OOXML part graph requires is registered: the slide itself,
   * its relationship part (linking the chosen layout), the
   * `[Content_Types].xml` override, the `<p:sldId>` entry and the presentation
   * relationship. Slides that already exist are left untouched, so appending
   * can never destroy a deck.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {number} [options.layoutIndex] - 0-based layout to build from.
   * @param {string} [options.layoutName] - layout name, e.g. "Пустой слайд".
   * @param {string} [options.layoutType] - PowerPoint layout type, e.g. "title".
   * @param {number} [options.baseSlideIndex] - clone an existing slide instead.
   * @param {string} [options.title]
   * @param {string} [options.subtitle]
   * @param {Array<string|object>} [options.paragraphs] - body text.
   * @param {number} [options.position] - 0-based insertion position (default: append).
   * @param {string} [options.outputPath]
   * @returns {Promise<object>}
   */
  async addSlide(filePath, options = {}) {
    const {
      outputPath = filePath,
      title = '',
      subtitle = '',
      paragraphs = null,
      position = null
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const presentationXml = zip.getText('ppt/presentation.xml')
    if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const descriptors = slideParts(zip)
    const layouts = slideLayouts(zip)
    const size = slideSize(zip)

    const names = zip.list()
    const newPart = nextNumber(names, /^ppt\/slides\/slide(\d+)\.xml$/, 1)

    let slideXml
    let layoutPath = null
    let layoutEntry = null

    if (options.baseSlideIndex !== undefined && options.baseSlideIndex !== null) {
      const base = descriptors[options.baseSlideIndex]
      if (!base) {
        throw new Error(`Cannot clone slide ${options.baseSlideIndex}: the deck has ${descriptors.length} slide(s)`)
      }
      slideXml = zip.getText(base.partPath)
      const baseLayoutRel = readRelationships(zip, base.partPath).find((r) => r.type === REL.slideLayout)
      layoutPath = baseLayoutRel ? baseLayoutRel.partPath : null
      layoutEntry = layouts.find((l) => l.partPath === layoutPath) || null
      // Cloned content must not carry the base slide's shape ids forward
      // unchanged: two objects with the same id break editing in R7.
      slideXml = reassignShapeIds(slideXml, highestShapeId(zip) + 1)
    } else {
      layoutEntry = this._resolveLayout(layouts, options)
      if (!layoutEntry) {
        throw new Error(
          'No slide layout available in this presentation. Pass baseSlideIndex to clone an existing slide.'
        )
      }
      layoutPath = layoutEntry.partPath
      slideXml = this._buildSlideFromLayout(zip, layoutEntry, { title, subtitle, paragraphs, size })
    }

    if (!options.baseSlideIndex && options.baseSlideIndex !== 0 && title && !options.paragraphs) {
      // `_buildSlideFromLayout` already placed the title.
    }

    // Relationship part: the layout link, plus nothing else. A cloned slide
    // must not inherit its base slide's notes-slide relationship, which belongs
    // to the base slide alone.
    const slideRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + (layoutPath
        ? `<Relationship Id="rId1" Type="${REL.slideLayout}" Target="${relativeTarget(`ppt/slides/slide${newPart}.xml`, layoutPath)}"/>`
        : '')
      + '</Relationships>'

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
      `<Relationship Id="${newRelId}" Type="${REL.slide}" Target="slides/slide${newPart}.xml"/></Relationships>`
    )

    // Content type override for the new slide part.
    const contentTypesPath = '[Content_Types].xml'
    const contentTypes = zip.getText(contentTypesPath)
    if (!contentTypes) throw new Error('Invalid PPTX: [Content_Types].xml not found')
    const updatedContentTypes = ensureOverride(
      contentTypes,
      `/ppt/slides/slide${newPart}.xml`,
      CONTENT_TYPE.slide
    )

    // <p:sldId> ids must be unique within the presentation.
    const sldIds = [...presentationXml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => Number(m[1]))
    const newSldId = sldIds.length > 0 ? Math.max(...sldIds) + 1 : 256
    const sldIdXml = `<p:sldId id="${newSldId}" r:id="${newRelId}"/>`
    const at = position === null || position === undefined
      ? null
      : clamp(position, 0, descriptors.length)
    const updatedPresentation = insertSldId(presentationXml, sldIdXml, at, descriptors)

    zip.setBuffer(`ppt/slides/slide${newPart}.xml`, Buffer.from(slideXml, 'utf8'))
    zip.setText(`ppt/slides/_rels/slide${newPart}.xml.rels`, slideRels)
    zip.setText(relsPath, relsXml)
    zip.setText(contentTypesPath, updatedContentTypes)
    zip.setText('ppt/presentation.xml', updatedPresentation)

    const ordered = extractElements(updatedPresentation, 'p:sldId')
    this._syncSlideCount(zip, ordered.length)

    const insertedAt = ordered.findIndex((el) => getAttribute(el.outerXml, 'id') === String(newSldId))
    await this._write(zip, outputPath)

    return {
      success: true,
      slideIndex: insertedAt,
      slideNumber: insertedAt + 1,
      partPath: `ppt/slides/slide${newPart}.xml`,
      layout: layoutEntry ? { name: layoutEntry.name, type: layoutEntry.type, index: layoutEntry.index } : null,
      slideCount: ordered.length,
      outputPath
    }
  }

  /**
   * Delete a slide, removing every part that exists only for it.
   *
   * Media referenced by the deleted slide is left in the package when another
   * part still uses it, and notes-slide parts are removed with their slide.
   *
   * @param {string} filePath
   * @param {number} slideIndex
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async deleteSlide(filePath, slideIndex, options = {}) {
    const { outputPath = filePath } = options
    const zip = await ZipArchive.fromFile(filePath)
    const presentationXml = zip.getText('ppt/presentation.xml')
    if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const descriptors = slideParts(zip)
    const target = descriptors[slideIndex]
    if (!target) {
      throw new Error(
        `Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s) (0..${descriptors.length - 1})`
      )
    }

    // Presentation entry and relationship.
    const sldIdElement = extractElements(presentationXml, 'p:sldId')
      .find((el) => getAttribute(el.outerXml, 'r:id') === target.relId)
    let updatedPresentation = presentationXml
    if (sldIdElement) updatedPresentation = presentationXml.replace(sldIdElement.outerXml, '')

    const relsPath = 'ppt/_rels/presentation.xml.rels'
    let relsXml = zip.getText(relsPath) || ''
    const slideRel = extractElements(relsXml, 'Relationship')
      .find((r) => getAttribute(r.outerXml, 'Id') === target.relId)
    if (slideRel) relsXml = relsXml.replace(slideRel.outerXml, '')

    // Notes slide belonging to the deleted slide.
    const notesRel = readRelationships(zip, target.partPath).find((r) => r.type === REL.notesSlide)

    const removed = [target.partPath, relsPathFor(target.partPath)]
    zip.remove(target.partPath)
    zip.remove(relsPathFor(target.partPath))
    if (notesRel && zip.has(notesRel.partPath)) {
      zip.remove(notesRel.partPath)
      zip.remove(relsPathFor(notesRel.partPath))
      removed.push(notesRel.partPath, relsPathFor(notesRel.partPath))
    }

    // Content type override points at a part that no longer exists.
    const contentTypesPath = '[Content_Types].xml'
    const contentTypes = zip.getText(contentTypesPath)
    if (contentTypes) {
      zip.setText(contentTypesPath, contentTypes.replace(
        new RegExp(`<Override PartName="/${escapeRegex(target.partPath)}"[^>]*/>`, 'g'),
        ''
      ))
    }

    zip.setText(relsPath, relsXml)
    zip.setText('ppt/presentation.xml', updatedPresentation)
    const remaining = extractElements(updatedPresentation, 'p:sldId').length
    this._syncSlideCount(zip, remaining)
    await this._write(zip, outputPath)

    return {
      success: true,
      deletedIndex: slideIndex,
      removedParts: removed,
      slideCount: remaining,
      outputPath
    }
  }

  /**
   * Duplicate a slide, inserting the copy directly after the original.
   * @param {string} filePath
   * @param {number} slideIndex
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async duplicateSlide(filePath, slideIndex, options = {}) {
    const { outputPath = filePath, position = null } = options
    const zip = await ZipArchive.fromFile(filePath)
    const presentationXml = zip.getText('ppt/presentation.xml')
    if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const descriptors = slideParts(zip)
    const source = descriptors[slideIndex]
    if (!source) {
      throw new Error(
        `Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s) (0..${descriptors.length - 1})`
      )
    }

    const names = zip.list()
    const newPart = nextNumber(names, /^ppt\/slides\/slide(\d+)\.xml$/, 1)
    // Duplicated objects must get fresh shape ids, or R7 sees two shapes with
    // the same id in one deck and refuses to edit either of them.
    const slideXml = reassignShapeIds(zip.getText(source.partPath), highestShapeId(zip) + 1)

    const sourceRels = readRelationships(zip, source.partPath)
    const relsPath = relsPathFor(source.partPath)
    let relsXml = zip.getText(relsPath)
    if (!relsXml) {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    }
    // The copy links the same layout but keeps no notes slide: a notes part
    // cannot belong to two slides.
    const kept = readRelationships(zip, source.partPath).filter((r) => r.type !== REL.notesSlide)
    const slideRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + kept.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.targetMode ? ` TargetMode="${r.targetMode}"` : ''}/>`).join('')
      + '</Relationships>'

    const presentationRelsPath = 'ppt/_rels/presentation.xml.rels'
    let presentationRels = zip.getText(presentationRelsPath) || ''
    const relIds = [...presentationRels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
    const newRelId = `rId${relIds.length > 0 ? Math.max(...relIds) + 1 : 1}`
    presentationRels = presentationRels.replace(
      '</Relationships>',
      `<Relationship Id="${newRelId}" Type="${REL.slide}" Target="slides/slide${newPart}.xml"/></Relationships>`
    )

    const contentTypesPath = '[Content_Types].xml'
    const contentTypes = zip.getText(contentTypesPath)
    if (!contentTypes) throw new Error('Invalid PPTX: [Content_Types].xml not found')

    const sldIds = [...presentationXml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => Number(m[1]))
    const newSldId = sldIds.length > 0 ? Math.max(...sldIds) + 1 : 256
    const sldIdXml = `<p:sldId id="${newSldId}" r:id="${newRelId}"/>`
    const insertAt = position === null || position === undefined ? slideIndex + 1 : clamp(position, 0, descriptors.length)
    const updatedPresentation = insertSldId(presentationXml, sldIdXml, insertAt, descriptors)

    zip.setBuffer(`ppt/slides/slide${newPart}.xml`, Buffer.from(slideXml, 'utf8'))
    zip.setText(`ppt/slides/_rels/slide${newPart}.xml.rels`, slideRels)
    zip.setText(presentationRelsPath, presentationRels)
    zip.setText(contentTypesPath, ensureOverride(contentTypes, `/ppt/slides/slide${newPart}.xml`, CONTENT_TYPE.slide))
    zip.setText('ppt/presentation.xml', updatedPresentation)

    const ordered = extractElements(updatedPresentation, 'p:sldId')
    this._syncSlideCount(zip, ordered.length)
    await this._write(zip, outputPath)

    const at = ordered.findIndex((el) => getAttribute(el.outerXml, 'id') === String(newSldId))
    return {
      success: true,
      sourceIndex: slideIndex,
      slideIndex: at,
      slideNumber: at + 1,
      partPath: `ppt/slides/slide${newPart}.xml`,
      slideCount: ordered.length,
      outputPath
    }
  }

  /**
   * Move a slide to another position in the deck.
   *
   * Only the `<p:sldId>` order changes: no slide part is rewritten, so nothing
   * on any slide can be affected by a reorder.
   *
   * @param {string} filePath
   * @param {number} fromIndex
   * @param {number} toIndex
   * @param {object} [options]
   * @returns {Promise<object>}
   */
  async moveSlide(filePath, fromIndex, toIndex, options = {}) {
    const { outputPath = filePath } = options
    const order = await this.reorderSlides(filePath, { order: null, fromIndex, toIndex, outputPath })
    return order
  }

  /**
   * Reorder the deck.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {number[]} [options.order] - the complete new order, as a permutation
   *   of 0..slideCount-1.
   * @param {number} [options.fromIndex] - single-slide move.
   * @param {number} [options.toIndex]
   * @param {string} [options.outputPath]
   * @returns {Promise<object>}
   */
  async reorderSlides(filePath, options = {}) {
    const { outputPath = filePath, order = null, fromIndex = null, toIndex = null } = options
    const zip = await ZipArchive.fromFile(filePath)
    const presentationXml = zip.getText('ppt/presentation.xml')
    if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

    const sldIdElements = extractElements(presentationXml, 'p:sldId')
    const count = sldIdElements.length

    let sequence
    if (Array.isArray(order)) {
      if (order.length !== count) {
        throw new Error(`order must list all ${count} slides, got ${order.length}`)
      }
      const seen = new Set(order)
      if (seen.size !== count || order.some((i) => !Number.isInteger(i) || i < 0 || i >= count)) {
        throw new Error(`order must be a permutation of 0..${count - 1}`)
      }
      sequence = order.map((i) => sldIdElements[i])
    } else {
      if (!Number.isInteger(fromIndex) || !Number.isInteger(toIndex)) {
        throw new Error('reorderSlides needs either order[], or integer fromIndex and toIndex')
      }
      if (fromIndex < 0 || fromIndex >= count) {
        throw new Error(`fromIndex ${fromIndex} is out of range 0..${count - 1}`)
      }
      const target = clamp(toIndex, 0, count - 1)
      const list = [...sldIdElements]
      const [moved] = list.splice(fromIndex, 1)
      list.splice(target, 0, moved)
      sequence = list
    }

    const rebuilt = presentationXml.replace(
      /(<p:sldIdLst>)([\s\S]*?)(<\/p:sldIdLst>)/,
      `$1${sequence.map((el) => el.outerXml).join('')}$3`
    )

    zip.setText('ppt/presentation.xml', rebuilt)
    await this._write(zip, outputPath)

    return {
      success: true,
      slideCount: count,
      order: sequence.map((el) => {
        const relId = getAttribute(el.outerXml, 'r:id')
        const descriptor = slideParts(zip).find((d) => d.relId === relId)
        return descriptor ? descriptor.partPath : null
      }),
      outputPath
    }
  }

  // ======================================================= object creation

  /**
   * Add a shape to a slide.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {number|string} options.shape - a name such as `rectangle`,
   *   `rounded-rectangle`, `ellipse`, `line`, `arrow`, `triangle`, or any
   *   DrawingML preset name.
   * @param {number} [options.slideIndex=0]
   * @param {string} [options.name]
   * @param {number} [options.zOrder] - insertion position in the shape tree.
   * @returns {Promise<object>}
   */
  async addShape(filePath, options = {}) {
    return this._addObject(filePath, 'shape', options)
  }

  /** Add a text box to a slide. Shorthand for `addShape` with a text body. */
  async addTextBox(filePath, options = {}) {
    return this._addObject(filePath, 'textBox', options)
  }

  /** Add a picture to a slide. */
  async addImage(filePath, options = {}) {
    return this._addObject(filePath, 'image', options)
  }

  async _addObject(filePath, kind, options = {}) {
    const { outputPath = filePath, slideIndex = 0, imagePath = null } = options

    if (kind === 'image' && !imagePath) {
      throw new Error('addImage requires imagePath')
    }

    const zip = await ZipArchive.fromFile(filePath)
    const descriptors = slideParts(zip)
    const descriptor = descriptors[slideIndex]
    if (!descriptor) {
      throw new Error(`Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s)`)
    }

    let slideXml = zip.getText(descriptor.partPath)
    const id = nextShapeId(slideXml)
    const size = slideSize(zip)

    let xml
    let image = null

    if (kind === 'image') {
      const loaded = loadImage(imagePath)
      const mediaPath = storeMedia(zip, loaded.buffer, loaded.extension)
      const relId = addImageRelationship(zip, descriptor.partPath, mediaPath)
      const box = fitImage(
        { width: loaded.width, height: loaded.height },
        options,
        size
      )
      xml = buildPicture({
        id,
        name: options.name,
        description: options.description,
        relId,
        ...box,
        rotation: options.rotation,
        flipHorizontal: options.flipHorizontal,
        flipVertical: options.flipVertical
      })
      image = {
        mediaPath,
        relId,
        format: loaded.format,
        naturalWidth: loaded.width,
        naturalHeight: loaded.height,
        bytes: loaded.bytes
      }
    } else {
      const preset = kind === 'textBox' ? 'rect' : resolvePreset(options.shape || 'rectangle')
      const paragraphs = kind === 'textBox'
        ? (options.paragraphs || (options.text !== undefined ? [options.text] : []))
        : (options.paragraphs || (options.text !== undefined ? [options.text] : null))
      xml = buildShape({
        id,
        name: options.name,
        description: options.description,
        preset,
        connector: preset === 'line' && options.connector !== false,
        ...options,
        paragraphs,
        // A text box has no fill and no outline unless the caller asks for one,
        // which is what "insert a text box" means in every editor.
        fill: options.fill !== undefined ? options.fill : (kind === 'textBox' ? null : undefined),
        line: options.line !== undefined ? options.line : (kind === 'textBox' ? null : undefined),
        noFill: options.noFill !== undefined ? options.noFill : (kind === 'textBox' && options.fill === undefined ? true : undefined)
      })
    }

    slideXml = insertIntoShapeTree(slideXml, xml, options.zOrder)
    zip.setText(descriptor.partPath, slideXml)
    await this._write(zip, outputPath)

    const result = {
      success: true,
      slideIndex,
      objectId: id,
      type: kind === 'image' ? 'image' : (options.shape || kind),
      outputPath
    }
    if (image) result.image = image
    return result
  }

  // ========================================================= object editing

  /**
   * Format one object: font, geometry, fill, stroke, alignment.
   *
   * Only the properties the caller names are changed. Everything else — the
   * gradient the author set, the shadow, the hyperlink, unknown extensions —
   * is left byte-identical.
   *
   * @param {string} filePath
   * @param {object} options
   * @param {number} [options.slideIndex=0]
   * @param {number} [options.objectId] - shape id, as reported by readSlide.
   * @param {string} [options.objectName] - shape name, as an alternative.
   * @param {number} [options.objectIndex] - position in the object list.
   * @param {Array<string|object>} [options.paragraphs] - replace the text body.
   * @param {string} [options.text] - replace the whole text with one paragraph.
   * @param {number} [options.paragraphIndex] - target one paragraph.
   * @returns {Promise<object>}
   */
  async formatObject(filePath, options = {}) {
    const { outputPath = filePath, slideIndex = 0 } = options
    const zip = await ZipArchive.fromFile(filePath)
    const { descriptor, slideXml, object } = this._locate(zip, slideIndex, options)

    // The object is located exactly once, against the slide as it is on disk.
    // Locating it *after* a rewrite cannot work: the markup the read model
    // cached no longer appears verbatim, and the old fall-through answered
    // "not found" by appending a brand-new shape for an object that was on the
    // slide all along. That is what put two `p:ph type="subTitle" idx="1"`
    // shapes and two `p:cNvPr id="3"` shapes on one slide.
    const located = locateElement(slideXml, object)
    const inserted = located.inserted
    let elementXml = located.elementXml

    if (options.paragraphs !== undefined || options.text !== undefined) {
      const paragraphs = options.paragraphs !== undefined ? options.paragraphs : [options.text]
      elementXml = replaceTextBody(elementXml, object, paragraphs, options)
    }

    if (options.imagePath) {
      const loaded = loadImage(options.imagePath)
      if (!object.image) throw new Error(`Object "${object.name || object.id}" is not an image`)
      const { partPath, replacedInPlace } = replaceMedia(zip, object.image, loaded)
      if (!replacedInPlace) {
        retargetImageRelationship(zip, descriptor.partPath, object.image.relId, partPath)
      }
      // A replacement with a different aspect ratio still has to fill the box
      // the caller asked for; an explicit size wins, otherwise the box stays.
      const box = {
        x: options.x === undefined ? object.x : options.x,
        y: options.y === undefined ? object.y : options.y,
        width: options.width === undefined ? object.width : options.width,
        height: options.height === undefined ? object.height : options.height
      }
      const patchedPicture = patchPicture(elementXml, {
        ...box,
        name: options.name,
        description: options.description
      })
      zip.setText(descriptor.partPath, commitElement(slideXml, located, patchedPicture))
      await this._write(zip, outputPath)
      return {
        success: true,
        slideIndex,
        objectId: object.id,
        image: { mediaPath: partPath, format: loaded.format, replacedInPlace },
        outputPath
      }
    }

    let patched = elementXml

    if (object.type === 'image') {
      patched = patchPicture(patched, {
        x: options.x,
        y: options.y,
        width: options.width,
        height: options.height,
        name: options.name,
        description: options.description
      })
    } else {
      patched = patchShapeProperties(patched, options)
      if (options.verticalAnchor !== undefined) {
        patched = patchTextBodyAttribute(patched, 'anchor', normalizeAnchor(options.verticalAnchor))
      }
    }

    // With `paragraphIndex` the same options apply to that one paragraph,
    // paragraph properties included, which is what "make paragraph 2 bold and
    // centred" means. Without it, the run properties go to every run and the
    // paragraph properties to every paragraph.
    const font = pickRunProperties(options)
    const hasFont = Object.keys(font).length > 0
    const wantsParagraphIndex = options.paragraphIndex !== undefined && options.paragraphIndex !== null
    const paragraphSpec = titleSafeParagraphProperties(object, pickParagraphProperties(options))
    const applyParagraph = Object.keys(paragraphSpec).length > 0 || wantsParagraphIndex

    if (hasFont || applyParagraph) {
      if (!firstElement(patched, 'p:txBody')) {
        patched = insertTextBody(patched, buildTextBody([]))
      }
      if (wantsParagraphIndex) {
        patched = patchOneParagraph(patched, options.paragraphIndex, paragraphSpec, font)
      } else {
        if (hasFont) patched = patchTextBodyFont(patched, font)
        if (Object.keys(paragraphSpec).length > 0) {
          patched = patchTextBodyParagraphProperties(patched, paragraphSpec)
        }
      }
    }

    // A title keeps the box the layout gives it and lets the renderer fit the
    // text, so a long line can never run into what follows.
    if (isTitleObject(object)) patched = ensureNormAutofit(patched)

    zip.setText(descriptor.partPath, commitElement(slideXml, located, patched))
    await this._write(zip, outputPath)

    return {
      success: true,
      slideIndex,
      objectId: object.id,
      name: object.name,
      materialised: inserted,
      changed: {
        font: hasFont ? font : undefined,
        paragraph: applyParagraph ? { ...paragraphSpec, paragraphIndex: options.paragraphIndex } : undefined,
        geometry: ['x', 'y', 'width', 'height', 'rotation'].some((k) => options[k] !== undefined) || undefined
      },
      outputPath
    }
  }

  /**
   * Replace the text of a placeholder or a named object.
   *
   * @param {string} filePath
   * @param {object} options - `objectId`/`objectName`/`placeholderType` to pick
   *   the target, `paragraphs` or `text` for the new content. Font options are
   *   applied at the same time.
   * @returns {Promise<object>}
   */
  async setText(filePath, options = {}) {
    return this.formatObject(filePath, options)
  }

  /**
   * Replace text in slide shapes by matching search string or target title.
   *
   * Kept signature-compatible with the original engine, and extended with
   * `objectId` for addressing a single object.
   *
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
      title = null,
      objectId = null,
      objectName = null
    } = options

    const zip = await ZipArchive.fromFile(filePath)
    const descriptors = slideParts(zip)
    const descriptor = descriptors[slideIndex]
    if (!descriptor) {
      throw new Error(`Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s)`)
    }

    let slideXml = zip.getText(descriptor.partPath)
    let matchesCount = 0

    if (objectId !== null || objectName !== null) {
      const { object } = this._locate(zip, slideIndex, options)
      if (replace !== null && replace !== undefined) {
        if (object.onSlide === false) {
          throw new Error(
            `Object "${object.name || object.id}" exists only in the layout; put it on the slide before editing it`
          )
        }
        const { elementXml } = locateElement(slideXml, object)
        slideXml = slideXml.replace(elementXml, replaceElementText(elementXml, search, replace))
        matchesCount++
      }
      zip.setText(descriptor.partPath, slideXml)
      await this._write(zip, outputPath)
      return { success: true, matchesCount, outputPath }
    }

    if (title) {
      slideXml = this._setSlideTitle(slideXml, title)
      matchesCount++
    }

    if (search && replace !== null && replace !== undefined) {
      const tElements = extractElements(slideXml, 'a:t')
      for (const t of tElements) {
        const text = decodeText(t.outerXml)
        if (text.includes(search)) {
          const newText = text.replaceAll(search, replace)
          slideXml = slideXml.replace(t.outerXml, `<a:t>${escapeText(newText)}</a:t>`)
          matchesCount++
        }
      }
    }

    zip.setText(descriptor.partPath, slideXml)
    await this._write(zip, outputPath)
    return { success: true, matchesCount, outputPath }
  }

  /**
   * Remove an object from a slide.
   * @param {string} filePath
   * @param {object} options
   * @returns {Promise<object>}
   */
  async removeObject(filePath, options = {}) {
    const { outputPath = filePath, slideIndex = 0, keepMedia = false } = options
    const zip = await ZipArchive.fromFile(filePath)
    const { descriptor, slideXml, object } = this._locate(zip, slideIndex, options)

    // A layout-only placeholder has nothing to remove: deleting it would
    // require an empty `<p:sp>` plus a "hide this placeholder" flag, which is a
    // different operation from removing a shape the author drew.
    if (object.onSlide === false) {
      throw new Error(
        `Object "${object.name || object.id}" exists only in the layout and cannot be removed from the slide`
      )
    }

    const located = locateElement(slideXml, object)

    zip.setText(descriptor.partPath, slideXml.replace(located.elementXml, ''))

    if (object.image && object.image.relId && !keepMedia) {
      dropImageRelationship(zip, descriptor.partPath, object.image.relId)
    }

    await this._write(zip, outputPath)
    return { success: true, slideIndex, removedObjectId: object.id, outputPath }
  }

  // ============================================================ internals

  /**
   * Resolve a slide's part plus the read model of the object a call addresses.
   * @private
   */
  _locate(zip, slideIndex, options) {
    const descriptors = slideParts(zip)
    const descriptor = descriptors[slideIndex]
    if (!descriptor) {
      throw new Error(`Slide index ${slideIndex} is out of range: the deck has ${descriptors.length} slide(s)`)
    }
    const slideXml = zip.getText(descriptor.partPath)
    const model = readSlideModel(zip, descriptor, {
      layouts: slideLayouts(zip),
      size: slideSize(zip),
      includeInherited: true
    })
    const object = this._selectObject(model.objects, options)
    return { descriptor, slideXml, object, model }
  }

  /** Pick one object from a read model by id, name, index or placeholder. */
  _selectObject(objects, options) {
    const { objectId = null, objectName = null, objectIndex = null, placeholderType = null } = options
    if (objectId !== null && objectId !== undefined) {
      const found = objects.find((o) => o.id === Number(objectId))
      if (!found) throw new Error(`Object id ${objectId} not found on this slide`)
      return found
    }
    if (objectName !== null && objectName !== undefined) {
      const found = objects.find((o) => o.name === objectName)
      if (!found) throw new Error(`Object named "${objectName}" not found on this slide`)
      return found
    }
    if (objectIndex !== null && objectIndex !== undefined) {
      const onSlide = objects.filter((o) => o.onSlide)
      const found = onSlide[objectIndex]
      if (!found) {
        throw new Error(`Object index ${objectIndex} is out of range: the slide has ${onSlide.length} object(s)`)
      }
      return found
    }
    if (placeholderType) {
      const wanted = String(placeholderType).toLowerCase()
      const found = objects.find((o) => o.placeholder && o.placeholder.type.toLowerCase() === wanted)
      if (!found) throw new Error(`No placeholder of type "${placeholderType}" on this slide`)
      return found
    }
    const firstText = objects.find((o) => o.onSlide && o.text && o.text.trim() !== '')
    if (firstText) return firstText
    const first = objects.find((o) => o.onSlide)
    if (!first) throw new Error('This slide has no objects')
    return first
  }

  /** Pick the layout a new slide should use. */
  _resolveLayout(layouts, options) {
    if (layouts.length === 0) return null
    const { layoutIndex = null, layoutName = null, layoutType = null, baseSlideIndex = null } = options

    if (layoutIndex !== null && layoutIndex !== undefined) {
      const found = layouts.find((l) => l.index === Number(layoutIndex))
      if (!found) {
        throw new Error(
          `Layout index ${layoutIndex} not found. Available: ${layouts.map((l) => `${l.index}=${l.name}`).join(', ')}`
        )
      }
      return found
    }
    if (layoutName) {
      const wanted = String(layoutName).toLowerCase()
      const exact = layouts.find((l) => l.name && l.name.toLowerCase() === wanted)
      if (exact) return exact
      const partial = layouts.find((l) => l.name && l.name.toLowerCase().includes(wanted))
      if (partial) return partial
      throw new Error(
        `No layout named "${layoutName}". Available: ${layouts.map((l) => l.name).join(', ')}`
      )
    }
    if (layoutType) {
      const wanted = String(layoutType).toLowerCase()
      const found = layouts.find((l) => String(l.type).toLowerCase() === wanted)
      if (!found) {
        throw new Error(
          `No layout of type "${layoutType}". Available types: ${[...new Set(layouts.map((l) => l.type))].join(', ')}`
        )
      }
      return found
    }
    if (baseSlideIndex !== null && baseSlideIndex !== undefined) {
      return null
    }
    // Sensible default: the first "title and content"-style layout, which is
    // what "add a slide" produces in every editor.
    const preferred = layouts.find((l) => l.type === 'obj')
      || layouts.find((l) => l.type === 'titleOnly')
      || layouts.find((l) => l.type !== 'title')
      || layouts[0]
    return preferred
  }

  /**
   * Build a slide part for a layout.
   *
   * The slide carries `<p:sp>` shells for the layout's content placeholders
   * only. Date, footer and slide-number fields are deliberately left to the
   * layout: writing them onto the slide freezes today's date into the deck and
   * stops the field from updating. What is written stays a placeholder, so its
   * geometry and typography keep coming from the layout and the master.
   */
  _buildSlideFromLayout(zip, layoutEntry, { title, subtitle, paragraphs, size }) {
    const layoutXml = zip.getText(layoutEntry.partPath)
    const definitions = layoutXml ? layoutDefinitions(layoutXml) : []
    const use = definitions.length > 0 ? definitions : defaultPlaceholders()

    const interactive = use.filter((p) => !DECORATIVE_PLACEHOLDERS.has(p.type))
    const chosen = interactive.length > 0 ? interactive : use

    let id = 1
    const shapes = []

    for (const definition of chosen) {
      id++
      const isTitle = definition.type === 'title' || definition.type === 'ctrTitle'
      const isSubtitle = definition.type === 'subTitle'
      let paragraphsForShape = []

      if (isTitle && title) {
        paragraphsForShape = [{ text: title }]
      } else if (isSubtitle && subtitle) {
        paragraphsForShape = [{ text: subtitle }]
      } else if (!isTitle && !isSubtitle && Array.isArray(paragraphs) && paragraphs.length > 0) {
        paragraphsForShape = paragraphs
      }

      shapes.push(buildShape({
        id,
        name: definition.name || `${definition.type} ${id}`,
        preset: 'rect',
        // The placeholder's own transform is deliberately *not* written: a
        // slide-level placeholder inherits its position from the layout, and
        // freezing that geometry on the slide would cut the slide loose from
        // the master.
        paragraphs: paragraphsForShape,
        placeholder: { type: definition.type, idx: definition.idx },
        style: null,
        verticalAnchor: definition.type === 'title' || definition.type === 'ctrTitle' ? 'b' : undefined,
        // A title that inherits its box has to fit its text to that box, or a
        // long line is drawn past the bottom and over the next object.
        autofit: isTitle ? 'shrink' : undefined
      }))
    }

    const width = size ? size.width : 12192000
    const height = size ? size.height : 6858000
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
      + '<p:cSld><p:spTree>'
      + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
      + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
      + `<a:ext cx="${width}" cy="${height}"/>`
      + `<a:chOff x="0" y="0"/><a:chExt cx="${width}" cy="${height}"/></a:xfrm></p:grpSpPr>`
      + shapes.join('')
      + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
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

  /** Write the package, creating the target directory when needed. */
  async _write(zip, outputPath, result = null) {
    const dir = path.dirname(outputPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    await zip.save(outputPath)
    return result
  }

  _setSlideTitle(slideXml, title) {
    return setFirstText(slideXml, title)
  }

  // ============================================================ validation

  /**
   * Validate PPTX integrity.
   * @param {string} filePath
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[], details: object}>}
   */
  async validate(filePath) {
    const errors = []
    const warnings = []
    let zip
    try {
      zip = await ZipArchive.fromFile(filePath)
    } catch (err) {
      return { valid: false, errors: [`ZIP corruption: ${err.message}`], warnings, details: {} }
    }

    if (!zip.has('[Content_Types].xml')) errors.push('Missing [Content_Types].xml')
    if (!zip.has('ppt/presentation.xml')) errors.push('Missing ppt/presentation.xml')
    if (!zip.has('ppt/slides/slide1.xml')) warnings.push('No ppt/slides/slide1.xml: the deck may use other part numbers')

    let descriptors = []
    try {
      descriptors = slideParts(zip)
    } catch (err) {
      errors.push(`Cannot resolve the slide list: ${err.message}`)
    }

    // Every slide the presentation lists must exist, and must link a layout
    // that also exists. A dangling relationship is the classic symptom of a
    // hand-built slide, and it is what makes R7 open a deck read-only.
    for (const descriptor of descriptors) {
      if (!zip.has(descriptor.partPath)) {
        errors.push(`Presentation lists ${descriptor.partPath}, which is not in the package`)
        continue
      }
      const rels = readRelationships(zip, descriptor.partPath)
      if (!rels.some((r) => r.type === REL.slideLayout)) {
        warnings.push(`${descriptor.partPath} has no slideLayout relationship`)
      }
      for (const rel of rels) {
        if (rel.targetMode === 'External') continue
        if (!zip.has(rel.partPath)) {
          errors.push(`${descriptor.partPath} references missing part ${rel.partPath}`)
        }
      }
    }

    for (const descriptor of descriptors) {
      if (!descriptor.partPath.endsWith('.xml')) continue
      if (!zip.has(descriptor.relsPath) && descriptors.length > 0) {
        warnings.push(`${descriptor.partPath} has no relationship part`)
      }
    }

    // Media parts must be announced in [Content_Types].xml.
    const contentTypes = zip.getText('[Content_Types].xml') || ''
    for (const media of listMedia(zip)) {
      const ext = media.partPath.split('.').pop()
      if (!new RegExp(`Extension="${ext}"`, 'i').test(contentTypes)) {
        errors.push(`Media part ${media.partPath} has no content-type Default for ".${ext}"`)
      }
    }

    // Structure first: a duplicated shape id or a placeholder claimed twice
    // makes a renderer draw one of the two and clip the other, which no other
    // check in here would notice.
    const structure = structuralReport(zip)
    for (const error of structure.errors) errors.push(error)

    return {
      valid: errors.length === 0,
      errors,
      warnings,
      details: {
        fileSize: fs.statSync(filePath).size,
        entries: zip.list().length,
        slides: descriptors.length,
        layouts: slideLayouts(zip).length,
        media: listMedia(zip).length,
        structure
      }
    }
  }

  /**
   * Validate the *structure* of every slide.
   *
   * OOXML requires `p:cNvPr/@id` to be unique inside a slide, forbids id 0,
   * and forbids two shapes claiming the same placeholder (`p:ph type` + `idx`).
   * A deck that breaks any of these opens without complaint and renders wrong:
   * the renderer draws whichever of the duplicates it meets first and clips the
   * other, which reaches the user as "the text is on top of the title".
   *
   * @param {string} filePath
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[], details: object}>}
   */
  async validateStructure(filePath) {
    let zip
    try {
      zip = await ZipArchive.fromFile(filePath)
    } catch (err) {
      return {
        valid: false,
        errors: [`ZIP corruption: ${err.message}`],
        warnings: [],
        details: { slides: [] }
      }
    }
    return structuralReport(zip)
  }

  /** Render the deck to PDF through the R7 converter. */
  async toPdf(filePath, outputPath) {
    const target = outputPath || filePath.replace(/\.pptx$/i, '.pdf')
    const result = await this.r7Adapter.convert(filePath, target)
    const head = fs.readFileSync(target).subarray(0, 8).toString('latin1')
    return { ...result, target, isPdf: head.startsWith('%PDF-') }
  }

  /**
   * The public shape catalogue, so a caller never has to guess a name.
   * @returns {object}
   */
  shapeCatalog() {
    return {
      shapes: Object.keys(SHAPE_PRESETS),
      presets: [...new Set(Object.values(SHAPE_PRESETS))].sort()
    }
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
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
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

// ============================================================ module helpers

function decodeText(xml) {
  return String(xml)
    .replace(/<a:br\s*\/>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .trim()
}

function escapeText(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function clamp(value, min, max) {
  return Math.min(Math.max(Number(value), min), max)
}

/**
 * Placeholders the layout renders on every slide.
 *
 * A date, footer or slide-number field belongs to the layout and the master:
 * writing one onto a slide freezes the current value into the deck and stops
 * the field updating, which is data loss dressed up as completeness.
 */
const DECORATIVE_PLACEHOLDERS = new Set(['dt', 'ftr', 'sldNum', 'hdr'])

/** The placeholder types that are the slide's title. */
const TITLE_PLACEHOLDERS = new Set(['title', 'ctrTitle'])

/** True when an object is the slide's title placeholder. */
function isTitleObject(object) {
  const type = object && object.placeholder ? object.placeholder.type : null
  return Boolean(type && TITLE_PLACEHOLDERS.has(String(type).toLowerCase()))
}

/** Paragraph properties a title must not be given, because its box is fixed. */
const TITLE_UNSAFE_PARAGRAPH_KEYS = ['lineSpacing', 'spaceBefore', 'spaceAfter']

/**
 * Drop the paragraph spacing a title cannot afford.
 *
 * A title placeholder inherits its box from the layout and never grows, so an
 * explicit line spacing or a space-before/after is drawn straight out of the
 * box and on top of the next object. Body lists are where those options belong;
 * a per-paragraph request inside `paragraphs` is still honoured verbatim.
 */
function titleSafeParagraphProperties(object, spec) {
  if (!isTitleObject(object)) return spec
  const out = { ...spec }
  for (const key of TITLE_UNSAFE_PARAGRAPH_KEYS) delete out[key]
  return out
}

/** The next free `slideN.xml` number. */
function nextNumber(names, regex, fallback) {
  const used = []
  for (const name of names) {
    const match = name.match(regex)
    if (match) used.push(Number(match[1]))
  }
  return used.length > 0 ? Math.max(...used) + 1 : fallback
}

/** Rewrite every `p:cNvPr/@id` so a copied slide has no duplicate ids. */
function reassignShapeIds(slideXml, startAt) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) return slideXml
  const elements = collectShapeElements(tree.innerXml)
  let next = Math.max(Number(startAt) || 0, 2)
  let out = slideXml

  for (const element of elements) {
    const cNvPr = firstElement(element.xml, 'p:cNvPr')
    if (!cNvPr) continue
    const id = Number(getAttribute(openingTag(cNvPr), 'id') || 0)
    // The replacement has to carry a *different* id. Reusing the original's
    // number makes the search string and the replacement identical, so the
    // copy silently keeps the ids of the slide it came from and R7 then
    // refuses to edit either of them.
    const renamed = cNvPr.replace(/id="\d+"/, `id="${next}"`)
    if (renamed !== cNvPr) out = out.replace(element.xml, element.xml.replace(cNvPr, renamed))
    next = Math.max(next + 1, id + 1)
  }
  return out
}

/**
 * The highest shape id in the whole package.
 *
 * Shape ids only have to be unique per slide, but a duplicated slide that
 * happens to start numbering where another slide already is makes the deck
 * awkward to address, so copies are numbered from the deck-wide maximum.
 */
function highestShapeId(zip) {
  let highest = 1
  for (const descriptor of slideParts(zip)) {
    const slideXml = zip.getText(descriptor.partPath)
    if (!slideXml) continue
    for (const match of slideXml.matchAll(/<p:cNvPr id="(\d+)"/g)) {
      const id = Number(match[1])
      if (id > highest) highest = id
    }
  }
  return highest
}

/** The next free shape id inside one slide. */
function nextShapeId(slideXml) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) return 2
  let max = 1
  for (const element of collectShapeElements(tree.innerXml)) {
    const cNvPr = firstElement(element.xml, 'p:cNvPr')
    if (!cNvPr) continue
    const id = Number(getAttribute(openingTag(cNvPr), 'id') || 0)
    if (id > max) max = id
  }
  return max + 1
}

/**
 * Check every slide's shape tree against the rules OOXML actually enforces.
 *
 *  - `p:cNvPr/@id` is unique inside a slide.
 *  - a shape does not reuse the shape tree group's (`p:nvGrpSpPr`) id.
 *  - no shape carries id 0, which is not a legal shape id.
 *  - no two shapes claim the same placeholder (`p:ph type` + `idx`).
 *
 * None of these stop a deck from opening, which is exactly why they need
 * checking: the renderer draws one of the duplicates, clips the other, and the
 * author sees two objects on top of each other.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @returns {{valid: boolean, errors: string[], warnings: string[], details: object}}
 */
function structuralReport(zip) {
  const errors = []
  const warnings = []
  const slides = []

  let descriptors = []
  try {
    descriptors = slideParts(zip)
  } catch (err) {
    errors.push(`Cannot resolve the slide list: ${err.message}`)
  }

  for (const descriptor of descriptors) {
    const label = descriptor.partPath.replace(/^ppt\/slides\//, '')
    const slideXml = zip.getText(descriptor.partPath) || ''
    const tree = extractElements(slideXml, 'p:spTree')[0]
    const report = {
      index: descriptor.index,
      slideNumber: descriptor.index + 1,
      partPath: descriptor.partPath,
      shapeCount: 0,
      groupId: null,
      duplicateShapeIds: [],
      duplicatePlaceholders: [],
      zeroIds: [],
      groupIdCollisions: []
    }
    slides.push(report)

    if (!tree) {
      errors.push(`${label}: has no <p:spTree>`)
      continue
    }

    const groupCnvPr = firstElement(firstElement(tree.innerXml, 'p:nvGrpSpPr') || '', 'p:cNvPr')
    report.groupId = groupCnvPr ? Number(getAttribute(groupCnvPr, 'id') || 0) : null

    const elements = collectShapeElements(tree.innerXml)
    report.shapeCount = elements.length

    const byId = new Map()
    const byPlaceholder = new Map()

    for (const element of elements) {
      const { id, name } = shapeIdentity(element.xml)
      const shapeLabel = name || `#${id}`
      if (id === 0) report.zeroIds.push(shapeLabel)
      if (report.groupId !== null && id === report.groupId) report.groupIdCollisions.push(shapeLabel)
      if (!byId.has(id)) byId.set(id, [])
      byId.get(id).push(shapeLabel)

      const placeholder = placeholderIdentity(element.xml)
      if (placeholder) {
        const key = `${placeholder.type}/${placeholder.idx === null ? '' : placeholder.idx}`
        if (!byPlaceholder.has(key)) byPlaceholder.set(key, [])
        byPlaceholder.get(key).push(shapeLabel)
      }
    }

    for (const [id, names] of byId) {
      if (names.length < 2) continue
      report.duplicateShapeIds.push({ id, names })
      errors.push(`${label}: duplicate shape id ${id} (${names.join(', ')})`)
    }
    for (const [key, names] of byPlaceholder) {
      if (names.length < 2) continue
      report.duplicatePlaceholders.push({ key, names })
      errors.push(`${label}: duplicate placeholder ${key} (${names.join(', ')})`)
    }
    for (const name of report.zeroIds) {
      errors.push(`${label}: shape "${name}" has id 0, which is not a legal shape id`)
    }
    for (const name of report.groupIdCollisions) {
      errors.push(`${label}: shape "${name}" reuses the shape-tree group id ${report.groupId}`)
    }
  }

  return { valid: errors.length === 0, errors, warnings, details: { slides } }
}

/** Append (or insert at a z-order position) an object into a slide's shape tree. */
function insertIntoShapeTree(slideXml, objectXml, zOrder = null) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) throw new Error('Invalid slide: <p:spTree> not found')

  if (zOrder === null || zOrder === undefined) {
    return slideXml.replace(tree.outerXml, tree.outerXml.replace('</p:spTree>', `${objectXml}</p:spTree>`))
  }

  const elements = collectShapeElements(tree.innerXml)
  const index = clamp(zOrder, 0, elements.length)
  if (index >= elements.length) {
    return slideXml.replace(tree.outerXml, tree.outerXml.replace('</p:spTree>', `${objectXml}</p:spTree>`))
  }
  const anchor = elements[index]
  const insertAt = anchor.index
  const newInner = tree.innerXml.slice(0, insertAt) + objectXml + tree.innerXml.slice(insertAt)
  return slideXml.replace(tree.innerXml, newInner)
}

/**
 * Locate an object's markup inside a slide.
 *
 * The read model caches the exact XML it parsed, so the first attempt is a
 * plain substring search — but that cache goes stale the moment anything
 * rewrites the object, and an object that exists only in the layout has no
 * markup on the slide at all. Both cases used to end in the same place: a
 * brand-new shape was appended for an object that was already there, which put
 * two shapes with one `cNvPr/@id` and two shapes with one `p:ph` on the slide.
 *
 * The search therefore falls back to identity, in the order a caller means:
 * the id it read, then the name, then the placeholder `(type, idx)`. Only a
 * layout-only object — one the read model reported with `onSlide: false` — is
 * materialised, and that copy always takes a fresh id.
 *
 * @param {string} slideXml
 * @param {object} object - a read-model object.
 * @returns {{elementXml: string, inserted: boolean}} `inserted` is true only
 *   when a real shape had to be placed on the slide for the edit to land.
 */
function locateElement(slideXml, object) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) throw new Error('Invalid slide: <p:spTree> not found')

  const elementXml = object._element
  if (elementXml && slideXml.includes(elementXml)) {
    return { elementXml, inserted: false }
  }

  const elements = collectShapeElements(tree.innerXml)

  if (object.id) {
    const found = elements.find((el) => shapeIdentity(el.xml).id === object.id)
    if (found) return { elementXml: found.xml, inserted: false }
  }

  if (object.name) {
    const found = elements.find((el) => shapeIdentity(el.xml).name === object.name)
    if (found) return { elementXml: found.xml, inserted: false }
  }

  if (object.placeholder) {
    const found = elements.find((el) => samePlaceholderOf(el.xml, object.placeholder))
    if (found) return { elementXml: found.xml, inserted: false }
  }

  // Nothing on the slide carries this object's identity. A layout-only
  // placeholder is the one case where a shape has to be written.
  if (object.onSlide === false) return materialisePlaceholder(slideXml, object)

  throw new Error(
    `Object "${object.name || object.id}" is reported on the slide but its markup was not found; `
    + 'refusing to add a second shape for it'
  )
}

/** Place a materialised object into the shape tree marked by an insertion. */
function withInsertedObject(slideXml, objectXml) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) throw new Error('Invalid slide: <p:spTree> not found')
  return slideXml.replace(tree.innerXml, `${tree.innerXml}${objectXml}`)
}

/** Write a located object back into its slide, appending it when it is new. */
function commitElement(slideXml, located, elementXml) {
  if (located.inserted) return withInsertedObject(slideXml, elementXml)
  return slideXml.replace(located.elementXml, elementXml)
}

/** The `cNvPr` identity of a shape element: its id and its name. */
function shapeIdentity(elementXml) {
  const cNvPr = openingTag(firstElement(elementXml, 'p:cNvPr') || '')
  return {
    id: Number(getAttribute(cNvPr, 'id') || 0),
    name: getAttribute(cNvPr, 'name') || null
  }
}

/** The `<p:ph>` descriptor of a shape element, as `(type, idx)`. */
function placeholderIdentity(elementXml) {
  const ph = firstElement(elementXml, 'p:ph')
  if (!ph) return null
  return {
    type: getAttribute(ph, 'type') || 'body',
    idx: getAttribute(ph, 'idx')
  }
}

/** True when a shape's placeholder is exactly the one an object addresses. */
function samePlaceholderOf(elementXml, placeholder) {
  const found = placeholderIdentity(elementXml)
  if (!found || !placeholder || !placeholder.type) return false
  if (found.type !== String(placeholder.type)) return false
  const wanted = placeholder.idx === null || placeholder.idx === undefined ? null : String(placeholder.idx)
  return found.idx === wanted
}

/**
 * Build a real slide-level shape from a layout placeholder definition.
 *
 * A placeholder the slide already carries is reused, never copied: the copy
 * would claim the same `p:ph` and leave the renderer free to draw either one.
 * A genuinely new shape takes its id from the slide, never from the layout —
 * the layout numbers its own shapes from 1 and the slide numbers from 2, so
 * inheriting the layout's id is what produced two shapes with id 3.
 *
 * @returns {{elementXml: string, inserted: boolean}}
 */
function materialisePlaceholder(slideXml, object) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) throw new Error('Invalid slide: <p:spTree> not found')

  if (object.placeholder) {
    const existing = collectShapeElements(tree.innerXml)
      .find((el) => samePlaceholderOf(el.xml, object.placeholder))
    if (existing) return { elementXml: existing.xml, inserted: false }
  }

  const placeholder = object.placeholder || {}

  // Paragraph shape is carried over, styling is not: a placeholder that had no
  // explicit formatting on the slide must keep inheriting it from the layout,
  // and freezing the resolved font onto the shape would break that link.
  const paragraphs = []
  for (const paragraph of object.paragraphs || []) {
    const rebuilt = { text: paragraph.text || '' }
    if (paragraph.level) rebuilt.level = paragraph.level
    if (paragraph.alignment) rebuilt.alignment = paragraph.alignment
    if (paragraph.bullet && paragraph.bullet !== 'inherit') {
      if (paragraph.bullet === 'none') rebuilt.bullet = false
      else if (paragraph.bullet === 'numbered') rebuilt.numbered = true
      else {
        rebuilt.bullet = paragraph.bulletCharacter || '•'
      }
    }
    if (Array.isArray(paragraph.runs) && paragraph.runs.length > 0) {
      rebuilt.runs = paragraph.runs.map((run) => ({ text: run.text }))
      delete rebuilt.text
    }
    paragraphs.push(rebuilt)
  }

  const elementXml = buildShape({
    id: nextShapeId(slideXml),
    name: object.name || `Placeholder ${placeholder.type}`,
    preset: object.preset && object.preset !== 'custom' ? object.preset : 'rect',
    placeholder: { type: placeholder.type, idx: placeholder.idx },
    // Geometry is inherited, so no transform is written. If the read model had
    // to fall back to a computed box, that box is written instead, because
    // otherwise the object would jump when it appears on the slide.
    x: object.inherited && object.x !== null ? object.x : undefined,
    y: object.inherited && object.y !== null ? object.y : undefined,
    width: object.inherited && object.width !== null ? object.width : undefined,
    height: object.inherited && object.height !== null ? object.height : undefined,
    paragraphs: paragraphs.length > 0 ? paragraphs : [],
    // A title shape that appears on the slide keeps the layout's box and lets
    // the renderer fit the text to it.
    autofit: isTitleObject(object) ? 'shrink' : undefined,
    style: null
  })

  return { elementXml, inserted: true }
}

/**
 * Replace the whole text body of an object.
 *
 * The object's own markup is passed in and returned: this is a pure rewrite of
 * one element, so the caller can decide where the result belongs. Locating the
 * element again after the rewrite is what used to create a duplicate.
 *
 * Only run-level and text-body options are forwarded: an `alignment` on the
 * object is a request about its paragraphs, not about every new paragraph the
 * caller happens to have left unaligned.
 */
function replaceTextBody(elementXml, object, paragraphs, options) {
  const bodySpec = {}
  for (const key of ['verticalAnchor', 'wrap', 'autofit', 'insetLeft', 'insetTop', 'insetRight', 'insetBottom']) {
    if (options[key] !== undefined) bodySpec[key] = options[key]
  }
  if (options.font && typeof options.font === 'object') bodySpec.font = options.font
  for (const key of ['family', 'complexFamily', 'size', 'bold', 'italic', 'underline', 'color', 'transparency', 'strike', 'caps', 'spacing', 'baseline', 'highlight']) {
    if (options[key] !== undefined) bodySpec[key] = options[key]
  }
  // A title keeps the box the layout gives it, so its text is fitted to that
  // box and never inflated by body-level paragraph spacing.
  if (isTitleObject(object)) {
    delete bodySpec.lineSpacing
    delete bodySpec.spaceBefore
    delete bodySpec.spaceAfter
    bodySpec.autofit = 'shrink'
  }

  const textBody = buildTextBody(paragraphs, bodySpec)
  const existing = firstElement(elementXml, 'p:txBody')

  if (existing) return elementXml.replace(existing, textBody)
  if (object.type === 'image') {
    throw new Error(`Object "${object.name || object.id}" is an image and cannot hold text`)
  }
  return insertTextBody(elementXml, textBody)
}

/** Insert a `<p:txBody>` into a shape that has none, in schema order. */
function insertTextBody(shapeXml, textBody) {
  if (firstElement(shapeXml, 'p:txBody')) {
    return shapeXml.replace(firstElement(shapeXml, 'p:txBody'), textBody)
  }
  const spPr = firstElement(shapeXml, 'p:spPr')
  if (spPr) return shapeXml.replace(spPr, `${spPr}${textBody}`)
  return shapeXml.replace(/(<p:nvSpPr>[\s\S]*?<\/p:nvSpPr>)/, `$1${textBody}`)
}

/** Replace the text of every `<a:t>` in the first text body of a slide. */
function setFirstText(slideXml, text) {
  const tree = extractElements(slideXml, 'p:spTree')[0]
  if (!tree) return slideXml
  const firstShape = collectShapeElements(tree.innerXml)[0]
  if (!firstShape) return slideXml
  const shapeXml = firstShape.xml
  const body = firstElement(shapeXml, 'p:txBody')
  const paragraph = `<a:p><a:r><a:rPr lang="ru-RU"/><a:t>${escapeText(text)}</a:t></a:r></a:p>`

  let patched
  if (body) {
    // Only the paragraphs are swapped. Writing a whole `<p:txBody>` here puts a
    // second `<a:bodyPr>` and a second `<a:lstStyle>` next to the ones the body
    // already has, and `p:txBody` allows exactly one of each — a schema
    // violation that a validator reports and a repair prompt acts on.
    const paragraphs = extractElements(body, 'a:p')
    const bodyXml = paragraphs.length === 0
      ? body.replace('</p:txBody>', `${paragraph}</p:txBody>`)
      : body.slice(0, paragraphs[0].index)
        + paragraph
        + body.slice(paragraphs[paragraphs.length - 1].index + paragraphs[paragraphs.length - 1].outerXml.length)
    patched = shapeXml.replace(body, bodyXml)
  } else {
    const textBody = `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraph}</p:txBody>`
    patched = shapeXml.replace(
      /(<p:spPr(?:\s[^>]*)?\/>|<p:spPr[\s\S]*?<\/p:spPr>)/,
      `$1${textBody}`
    )
  }

  // The title is fitted to the box it inherits, never allowed to grow past it.
  patched = ensureNormAutofit(patched)
  return slideXml.replace(shapeXml, patched)
}

/** Replace a search string inside one object's text. */
function replaceElementText(elementXml, search, replace) {
  if (!search) return elementXml
  let out = elementXml
  for (const t of extractElements(elementXml, 'a:t')) {
    const text = decodeText(t.outerXml)
    if (!text.includes(search)) continue
    out = out.replace(t.outerXml, `<a:t>${escapeText(text.replaceAll(search, replace))}</a:t>`)
  }
  return out
}

/** Placeholder definitions of a layout, restricted to the ones a slide uses. */
function layoutDefinitions(layoutXml) {
  const out = []
  for (const sp of extractElements(layoutXml, 'p:sp')) {
    const ph = firstElement(sp.outerXml, 'p:ph')
    if (!ph) continue
    const cNvPr = openingTag(firstElement(sp.outerXml, 'p:cNvPr') || '')
    out.push({
      type: getAttribute(ph, 'type') || 'body',
      idx: getAttribute(ph, 'idx'),
      name: getAttribute(cNvPr, 'name') || null
    })
  }
  return out
}

/** The placeholders a deck with no usable layout falls back to. */
function defaultPlaceholders() {
  return [
    { type: 'title', idx: null, name: 'Заголовок' },
    { type: 'body', idx: '1', name: 'Текст' }
  ]
}

/** Insert a `<p:sldId>` at a position, creating the list when it is missing. */
function insertSldId(presentationXml, sldIdXml, position, descriptors) {
  const list = extractElements(presentationXml, 'p:sldIdLst')[0]
  if (!list) {
    const built = `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`
    return /<\/p:sldMasterIdLst>/.test(presentationXml)
      ? presentationXml.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst>${built}`)
      : presentationXml.replace(/(<p:presentation\b[^>]*>)/, `$1${built}`)
  }

  const entries = extractElements(list.innerXml, 'p:sldId')
  const at = position === null || position === undefined
    ? entries.length
    : clamp(position, 0, entries.length)

  let newInner
  if (at >= entries.length) {
    newInner = `${list.innerXml}${sldIdXml}`
  } else {
    const anchor = entries[at]
    newInner = list.innerXml.slice(0, anchor.index) + sldIdXml + list.innerXml.slice(anchor.index)
  }
  return presentationXml.replace(list.innerXml, newInner)
}

/** Path of `targetPart` as seen from `sourcePart`, re-exported for addSlide. */
function relativeTarget(sourcePart, targetPart) {
  const from = sourcePart.split('/').slice(0, -1)
  const to = targetPart.split('/')
  let common = 0
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++
  return `${'../'.repeat(from.length - common)}${to.slice(common).join('/')}`
}

/** Pick the run-level properties from a mixed options object. */
function pickRunProperties(options) {
  const out = {}
  for (const key of ['family', 'complexFamily', 'size', 'bold', 'italic', 'underline', 'color', 'transparency', 'strike', 'caps', 'spacing', 'baseline', 'highlight']) {
    if (options[key] !== undefined && options[key] !== null) out[key] = options[key]
  }
  if (options.font && typeof options.font === 'object') {
    for (const [key, value] of Object.entries(options.font)) {
      if (value !== undefined && value !== null) out[key] = value
    }
  }
  if (options.fontColor !== undefined && options.color === undefined) out.color = options.fontColor
  return out
}

/** Pick the paragraph-level properties from a mixed options object. */
function pickParagraphProperties(options) {
  const out = {}
  for (const key of ['alignment', 'level', 'bullet', 'numbered', 'bulletType', 'bulletCharacter', 'lineSpacing', 'spaceBefore', 'spaceAfter', 'marginLeft', 'indent']) {
    if (options[key] !== undefined && options[key] !== null) out[key] = options[key]
  }
  if (options.paragraph && typeof options.paragraph === 'object') {
    for (const [key, value] of Object.entries(options.paragraph)) {
      if (value !== undefined && value !== null) out[key] = value
    }
  }
  return out
}
