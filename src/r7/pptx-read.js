import { extractElements, getAttribute } from '../shared/xml.js'
import {
  REL,
  argbToHex,
  childElement,
  findElement,
  firstElement,
  openingTag,
  readRelationships,
  rootOpeningTag,
  slideLayouts,
  slideSize,
  transparencyFromArgb,
  layoutPlaceholders
} from './pptx-util.js'
import { rotationToDegrees } from './pptx-build.js'

/**
 * The PPTX reader.
 *
 * Turns raw OOXML into the structure a model can reason about without ever
 * seeing a namespace prefix: slides, their objects, each object's geometry,
 * text, font, fill, stroke and paragraphs — including placeholders whose
 * geometry and typography live in the layout and the master rather than on the
 * slide itself.
 */

/** The `<p:spTree>` of a slide, or null when the part is not a slide. */
export function shapeTree(slideXml) {
  return findElement(slideXml, 'p:spTree')
}

/** The `<p:cSld>` name attribute of a layout, master or slide. */
function slideName(xml) {
  const cSld = openingTag(firstElement(xml, 'p:cSld') || '')
  return getAttribute(cSld, 'name') || null
}

/**
 * The typography a placeholder inherits, resolved down the master chain.
 *
 * For each level the sources are ordered the way PowerPoint resolves them:
 *
 *  1. the layout's own placeholder `<a:lstStyle>`,
 *  2. the layout placeholder's `<a:pPr>` default run properties,
 *  3. the master's matching placeholder `<a:lstStyle>`,
 *  4. the master's placeholder paragraph default run properties,
 *  5. the master's `<p:txStyles><p:titleStyle>` or `<p:bodyStyle>`.
 *
 * The first source that defines a property wins. Skipping step 5 is what makes
 * a title read back as the presentation's 18 pt default instead of the 44 pt
 * the master actually renders, which is the kind of wrong answer that leads a
 * model to "fix" a font that was never broken.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string|null} layoutPath
 * @param {{type: string, idx: string|null}|null} placeholder
 * @returns {{levels: Array<string|undefined>}}
 */
function inheritedPlaceholderStyle(zip, layoutPath, placeholder, chain = null) {
  const levels = []
  if (!placeholder) return { levels }
  const xmls = chain || masterChain(zip, layoutPath)

  const mergeLevel = (level, levelXml) => {
    if (level < 0 || level > 8 || !levelXml) return
    if (!levels[level]) {
      levels[level] = levelXml
      return
    }
    const incoming = firstElement(levelXml, 'a:defRPr')
    if (!incoming) return
    const existing = levels[level]
    if (firstElement(existing, 'a:defRPr')) return
    const tag = `<a:lvl${level + 1}pPr>`
    if (existing.includes(tag)) levels[level] = existing.replace(tag, `${tag}${incoming}`)
  }

  for (const xml of xmls.parts) {
    for (const sp of extractElements(xml, 'p:sp')) {
      const ph = placeholderOf(sp.outerXml)
      if (!ph) continue
      if (ph.type !== placeholder.type) continue
      if (placeholder.idx !== null && ph.idx !== null && ph.idx !== placeholder.idx) continue

      const lst = firstElement(sp.outerXml, 'a:lstStyle')
      if (lst) {
        const parsed = parseLevelStyles(lst)
        for (let level = 0; level < 9; level++) mergeLevel(level, parsed[level])
      }
      // The placeholder's sample paragraph carries the default run properties
      // the layout intends for that field even when it declares no list style.
      for (const p of extractElements(sp.outerXml, 'a:p')) {
        const pPr = firstElement(p.outerXml, 'a:pPr')
        if (!pPr) continue
        const level = Number(getAttribute(openingTag(pPr), 'lvl') || 0)
        const defRPr = firstElement(pPr, 'a:defRPr')
        if (!defRPr) continue
        mergeLevel(level, `<a:lvl${level + 1}pPr>${defRPr}</a:lvl${level + 1}pPr>`)
      }
    }
  }

  // The master's text styles are the last word before the presentation default.
  const textStyles = masterTextStyle(xmls.master, placeholder.type)
  if (textStyles) {
    const parsed = parseLevelStyles(textStyles)
    for (let level = 0; level < 9; level++) mergeLevel(level, parsed[level])
  }

  return { levels }
}

/**
 * The layout XML and its master, read once per slide.
 *
 * Placeholder typography resolution walks the same two parts for every object
 * on the slide, and re-reading relationship parts per object turns a single
 * slide read into hundreds of package lookups.
 */
function masterChain(zip, layoutPath) {
  const parts = []
  let master = null
  if (layoutPath && zip.has(layoutPath)) {
    parts.push(zip.getText(layoutPath))
    const masterRel = readRelationships(zip, layoutPath).find((r) => r.type === REL.slideMaster)
    if (masterRel && zip.has(masterRel.partPath)) {
      master = zip.getText(masterRel.partPath)
      parts.push(master)
    }
  }
  return { parts, master }
}

/** The master's `<p:titleStyle>` or `<p:bodyStyle>` for a placeholder type. */
function masterTextStyle(masterXml, placeholderType) {
  if (!masterXml) return null
  const txStyles = firstElement(masterXml, 'p:txStyles')
  if (!txStyles) return null
  const tag = placeholderType === 'title' || placeholderType === 'ctrTitle'
    ? 'p:titleStyle'
    : 'p:bodyStyle'
  return firstElement(txStyles, tag)
}

/** Placeholder descriptor of the first `<p:ph>` inside a shape, or null. */
function placeholderOf(shapeXml) {
  const ph = firstElement(shapeXml, 'p:ph')
  if (!ph) return null
  return {
    type: getAttribute(ph, 'type') || 'body',
    idx: getAttribute(ph, 'idx'),
    orient: getAttribute(ph, 'orient'),
    size: getAttribute(ph, 'sz')
  }
}

/** Key under which a placeholder's inherited definition is filed. */
function placeholderKey(placeholder) {
  return `${placeholder.type}|${placeholder.idx === null ? '' : placeholder.idx}`
}

/**
 * A placeholder matches a layout definition when its type and index agree.
 * PowerPoint treats a missing index as "the one placeholder of that type",
 * which is what makes a bare `<p:ph type="title"/>` find the layout's title.
 */
function matchPlaceholder(placeholder, definitions) {
  const byIndex = definitions.get(placeholderKey(placeholder))
  if (byIndex) return byIndex
  const candidates = [...definitions.values()].filter((d) => d.type === placeholder.type)
  if (candidates.length === 0) return null
  const withGeometry = candidates.find((d) => d.x !== null && d.y !== null)
  return withGeometry || candidates[0]
}

/**
 * Collect every geometry and typography definition a layout and its master
 * offer, so a slide's placeholders can report what they actually inherit.
 */
function inheritedDefinitions(zip, layoutPath) {
  const definitions = new Map()
  const textStyles = []
  let masterPath = null
  let themePath = null

  if (layoutPath && zip.has(layoutPath)) {
    const layoutXml = zip.getText(layoutPath)
    const masterRel = readRelationships(zip, layoutPath).find((r) => r.type === REL.slideMaster)
    const masterXml = masterRel && zip.has(masterRel.partPath) ? zip.getText(masterRel.partPath) : null
    const masterPlaceholders = new Map(
      layoutPlaceholders(masterXml || '').map((ph) => [placeholderKey(ph), ph])
    )

    for (const ph of layoutPlaceholders(layoutXml, masterPlaceholders)) {
      definitions.set(placeholderKey(ph), ph)
    }
    textStyles.push(...listStyles(layoutXml))
    if (masterXml) {
      masterPath = masterRel.partPath
      for (const ph of layoutPlaceholders(masterXml)) {
        const key = placeholderKey(ph)
        // The layout wins over the master for the same placeholder.
        if (!definitions.has(key)) definitions.set(key, ph)
      }
      textStyles.push(...listStyles(masterXml))
    }
  }

  if (masterPath) {
    const themeRel = readRelationships(zip, masterPath).find((r) => r.type === REL.theme)
    if (themeRel) themePath = themeRel.partPath
  }

  return { definitions, textStyles, masterPath, themePath }
}

/** Every `<a:lstStyle>` in a layout or master, tagged by which part it is. */
export function listStyles(xml) {
  const out = []
  const source = rootOpeningTag(xml, 'p:sldLayout') ? 'layout' : 'master'
  for (const sp of extractElements(xml, 'p:sp')) {
    const lst = firstElement(sp.outerXml, 'a:lstStyle')
    if (!lst) continue
    const ph = placeholderOf(sp.outerXml)
    out.push({
      source,
      placeholderType: ph ? ph.type : null,
      levels: parseLevelStyles(lst)
    })
  }
  return out
}

/** Parse `<a:lvlNpPr>` children of an `<a:lstStyle>` into a level array. */
function parseLevelStyles(lstStyleXml) {
  const levels = []
  for (const pPr of extractElements(lstStyleXml, 'a:lvl1pPr')
    .concat(extractElements(lstStyleXml, 'a:lvl2pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl3pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl4pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl5pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl6pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl7pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl8pPr'))
    .concat(extractElements(lstStyleXml, 'a:lvl9pPr'))) {
    const level = Number((openingTag(pPr.outerXml).match(/a:lvl(\d)pPr/) || [])[1] || 1) - 1
    levels[level] = pPr.outerXml
  }
  return levels
}

/**
 * Resolve the effective run style of a text body.
 *
 * The chain is: run properties, paragraph defaults, the layout/master list
 * style for this placeholder and level, then the presentation's own
 * `p:defaultTextStyle`. The first source that defines a property wins — the
 * same order PowerPoint resolves it in.
 */
function resolveRunStyle(bodyXml, paragraphXml, context) {
  const sources = []
  const run = firstElement(paragraphXml, 'a:r')
  if (run) {
    const rPr = firstElement(run, 'a:rPr')
    if (rPr) sources.push(rPr)
  }
  const endPara = firstElement(paragraphXml, 'a:endParaRPr')
  if (endPara) sources.push(endPara)

  const pPr = firstElement(paragraphXml, 'a:pPr')
  const level = pPr ? Number(getAttribute(openingTag(pPr), 'lvl') || 0) : 0
  if (pPr) {
    const defRPr = firstElement(pPr, 'a:defRPr')
    if (defRPr) sources.push(defRPr)
  }

  const bodyLstStyle = firstElement(bodyXml, 'a:lstStyle') || null
  if (bodyLstStyle) {
    const levels = parseLevelStyles(bodyLstStyle)
    if (levels[level]) {
      const defRPr = firstElement(levels[level], 'a:defRPr')
      if (defRPr) sources.push(defRPr)
    }
  }

  const inherited = context.placeholderStyle
  if (inherited && inherited.levels[level]) {
    const defRPr = firstElement(inherited.levels[level], 'a:defRPr')
    if (defRPr) sources.push(defRPr)
  }

  if (context.defaultTextStyle) {
    const defRPr = context.defaultTextStyle[level]
    if (defRPr) sources.push(defRPr)
  }

  const merge = (pick) => {
    for (const source of sources) {
      const value = pick(source)
      if (value !== null && value !== undefined) return value
    }
    return null
  }

  const latins = merge((s) => {
    const tag = firstElement(s, 'a:latin')
    return tag ? getAttribute(tag, 'typeface') : null
  })
  const size = merge((s) => {
    const value = getAttribute(openingTag(s), 'sz')
    return value === null ? null : Number(value) / 100
  })
  const bold = merge((s) => booleanAttr(openingTag(s), 'b'))
  const italic = merge((s) => booleanAttr(openingTag(s), 'i'))
  const underlineRaw = merge((s) => {
    const value = getAttribute(openingTag(s), 'u')
    if (value === null) return null
    return value === 'none' ? false : value
  })
  const strike = merge((s) => booleanAttr(openingTag(s), 'strike'))
  const colorInfo = merge((s) => colourOf(s))

  return {
    family: latins,
    size,
    bold: bold === null ? null : bold,
    italic: italic === null ? null : italic,
    underline: underlineRaw === null ? null : underlineRaw,
    strike,
    color: colorInfo ? colorInfo.hex : null,
    colorToken: colorInfo ? colorInfo.token : null,
    transparency: colorInfo ? colorInfo.transparency : 0
  }
}

/**
 * The colour of a properties element, as a literal or a theme token.
 *
 * `a:srgbClr/@val` is `ST_HexColorRGB` — exactly six hex digits — and opacity
 * belongs in a child `a:alpha` (thousandths of a percent). Files written before
 * that was corrected carry an eight-digit `AARRGGBB` in `val` instead, so both
 * forms are read here; the eight-digit one is a legacy shape, not a valid one.
 */
function colourOf(propertiesXml) {
  const fill = firstElement(propertiesXml, 'a:solidFill')
  if (!fill) return null
  const srgb = firstElement(fill, 'a:srgbClr')
  if (srgb) {
    const value = getAttribute(openingTag(srgb), 'val')
    if (!value) return null
    const alphaEl = firstElement(srgb, 'a:alpha')
    const transparency = alphaEl
      ? 1 - Number(getAttribute(openingTag(alphaEl), 'val') || 100000) / 100000
      : (value.length === 8 ? transparencyFromArgb(value) : 0)
    return {
      hex: argbToHex(value.length === 8 ? value : `FF${value}`),
      token: null,
      transparency: Math.round(Math.min(Math.max(transparency, 0), 1) * 1000) / 1000
    }
  }
  for (const [tag, prefix] of [['a:schemeClr', 'scheme:'], ['a:prstClr', 'preset:'], ['a:sysClr', 'system:']]) {
    const el = firstElement(fill, tag)
    if (el) {
      const value = getAttribute(openingTag(el), 'val')
      const alphaEl = firstElement(el, 'a:alpha')
      const transparency = alphaEl
        ? 1 - Number(getAttribute(openingTag(alphaEl), 'val') || 100000) / 100000
        : 0
      return {
        hex: null,
        token: `${prefix}${value}`,
        transparency: Math.round(Math.min(Math.max(transparency, 0), 1) * 1000) / 1000
      }
    }
  }
  return null
}

/** `@b="1"` and friends, as a boolean or null when unset. */
function booleanAttr(tagXml, name) {
  const value = getAttribute(tagXml, name)
  if (value === null) return null
  return value === '1' || value.toLowerCase() === 'true'
}

/** Parse the presentation's `<p:defaultTextStyle>` into per-level `<a:defRPr>`. */
export function presentationDefaultTextStyle(zip) {
  const xml = zip.getText('ppt/presentation.xml')
  if (!xml) return []
  const style = firstElement(xml, 'p:defaultTextStyle')
  if (!style) return []
  return parseLevelStyles(style).map((levelXml) => (levelXml ? firstElement(levelXml, 'a:defRPr') : null))
}

/** The pixel size of a media part, for reporting an image's natural size. */
export function imageInfo(zip, partPath) {
  const buffer = zip.getBuffer(partPath)
  if (!buffer) return null
  return { partPath, bytes: buffer.length, format: sniffImageFormat(buffer) }
}

/** Identify an image format from its magic bytes. */
export function sniffImageFormat(buffer) {
  if (!buffer || buffer.length < 4) return null
  if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) return 'png'
  if (buffer[0] === 0xff && buffer[1] === 0xd8) return 'jpeg'
  if (buffer.subarray(0, 3).toString('latin1') === 'GIF') return 'gif'
  if (buffer.subarray(0, 2).toString('latin1') === 'BM') return 'bmp'
  if (buffer.subarray(0, 4).toString('latin1') === 'RIFF'
    && buffer.subarray(8, 12).toString('latin1') === 'WEBP') return 'webp'
  if (buffer.subarray(0, 4).toString('latin1').startsWith('<?xm')) return 'svg'
  return null
}

/**
 * Read one slide into the normalized model.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {object} descriptor - an entry from `slideParts`.
 * @param {object} [options]
 * @param {boolean} [options.includeInherited=true] - also report placeholder
 *   objects that exist only in the layout. They are what the slide will show,
 *   and their geometry is reported with `inherited: true`.
 * @param {Array} [options.layouts] - pre-computed layout list, to avoid
 *   re-scanning the package once per slide.
 * @param {object} [options.size] - pre-computed slide size.
 * @returns {object}
 */
export function readSlide(zip, descriptor, options = {}) {
  const slideXml = zip.getText(descriptor.partPath)
  if (!slideXml) throw new Error(`Slide part not found: ${descriptor.partPath}`)

  const layouts = options.layouts || slideLayouts(zip)
  const size = options.size || slideSize(zip)
  const includeInherited = options.includeInherited !== false

  const slideRels = readRelationships(zip, descriptor.partPath)
  const layoutRel = slideRels.find((r) => r.type === REL.slideLayout)
  const layoutPath = layoutRel ? layoutRel.partPath : null
  const layoutXml = layoutPath && zip.has(layoutPath) ? zip.getText(layoutPath) : null
  const layoutEntry = layouts.find((l) => l.partPath === layoutPath) || null

  const inherited = inheritedDefinitions(zip, layoutPath)
  const styleChain = masterChain(zip, layoutPath)
  const styleCache = new Map()
  const placeholderStyleFor = (placeholder) => {
    if (!placeholder) return null
    const key = placeholderKey(placeholder)
    if (!styleCache.has(key)) {
      styleCache.set(key, inheritedPlaceholderStyle(zip, layoutPath, placeholder, styleChain))
    }
    return styleCache.get(key)
  }

  const context = {
    placeholder: null,
    placeholderStyle: null,
    defaultTextStyle: presentationDefaultTextStyle(zip)
  }

  const relTargets = new Map(slideRels.map((r) => [r.id, r]))
  const tree = shapeTree(slideXml)
  const objects = []
  const presentKeys = new Set()

  if (tree) {
    const elements = collectShapeElements(tree.innerXml)
    for (const [position, element] of elements.entries()) {
      const placeholder = placeholderOf(element.xml)
      if (placeholder) presentKeys.add(placeholderKey(placeholder))
      context.placeholder = placeholder
      context.placeholderStyle = placeholderStyleFor(placeholder)
      const definition = placeholder ? matchPlaceholder(placeholder, inherited.definitions) : null
      objects.push(describeObject({
        element,
        position,
        placeholder,
        definition,
        zip,
        relTargets,
        context,
        layoutPath,
        inheritedDefinitions: inherited.definitions,
        size
      }))
    }
  }

  // Geometry and typography a placeholder inherits from the layout: the slide
  // itself carries an empty `<p:sp>`, so without this the most important text
  // on the deck reads back with no position and no font.
  if (includeInherited && layoutXml) {
    for (const ph of layoutPlaceholders(layoutXml)) {
      if (presentKeys.has(placeholderKey(ph))) continue
      const placeholder = { type: ph.type, idx: ph.idx, orient: ph.orient, size: null }
      context.placeholder = placeholder
      context.placeholderStyle = placeholderStyleFor(placeholder)
      objects.push(describeObject({
        element: syntheticPlaceholder(ph),
        position: objects.length,
        placeholder,
        definition: ph,
        zip,
        relTargets,
        context,
        layoutPath,
        inheritedDefinitions: inherited.definitions,
        size,
        inherited: true,
        onSlide: false
      }))
    }
  }

  return {
    index: descriptor.index,
    slideNumber: descriptor.index + 1,
    slideId: descriptor.sldId,
    partPath: descriptor.partPath,
    name: slideName(slideXml),
    layout: {
      partPath: layoutPath,
      name: layoutEntry ? layoutEntry.name : (layoutXml ? layoutNameOf(layoutXml) : null),
      type: layoutEntry ? layoutEntry.type : (layoutXml ? (getAttribute(rootOpeningTag(layoutXml, 'p:sldLayout'), 'type') || null) : null),
      index: layoutEntry ? layoutEntry.index : null
    },
    size: { width: size.width, height: size.height },
    hasNotes: slideRels.some((r) => r.type === REL.notesSlide),
    objects
  }
}

function layoutNameOf(layoutXml) {
  const cSld = openingTag(firstElement(layoutXml, 'p:cSld') || '')
  return getAttribute(cSld, 'name') || null
}

/**
 * Every top-level object element of a shape tree, in document order.
 *
 * The first two children of a `<p:spTree>` are the non-visual group properties
 * and the group transform: they are not objects, but `extractElements` cannot
 * tell them apart from a `<p:sp>`, so they are filtered by name.
 */
export function collectShapeElements(treeInnerXml) {
  const found = []
  for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp', 'p:contentPart']) {
    for (const el of extractElements(treeInnerXml, tag)) {
      found.push({ tag, xml: el.outerXml, index: el.index })
    }
  }
  found.sort((a, b) => a.index - b.index)
  return found
}

/** Build a stand-in element for a layout placeholder the slide does not carry. */
function syntheticPlaceholder(ph) {
  return {
    tag: 'p:sp',
    xml: `<p:sp><p:nvSpPr><p:cNvPr id="0" name="${ph.name.replace(/"/g, '&quot;')}"/>`
      + '<p:cNvSpPr/><p:nvPr><p:ph'
      + (ph.type ? ` type="${ph.type}"` : '')
      + (ph.idx !== null && ph.idx !== undefined ? ` idx="${ph.idx}"` : '')
      + '/></p:nvPr></p:nvSpPr><p:spPr/>'
      + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></p:txBody></p:sp>'
  }
}

/** Describe one shape element as a read-model object. */
function describeObject({
  element, position, placeholder, definition, zip, relTargets, context,
  layoutPath, inheritedDefinitions: definitions, size, inherited = false, onSlide = true
}) {
  const xml = element.xml
  const cNvPr = openingTag(firstElement(xml, 'p:cNvPr') || '')
  const id = Number(getAttribute(cNvPr, 'id') || 0)
  const name = getAttribute(cNvPr, 'name') || null
  const description = getAttribute(cNvPr, 'descr') || null

  const spPr = firstElement(xml, 'p:spPr')
  const xfrm = spPr ? firstElement(spPr, 'a:xfrm') : null
  const off = xfrm ? firstElement(xfrm, 'a:off') : null
  const ext = xfrm ? firstElement(xfrm, 'a:ext') : null

  let x = off ? Number(getAttribute(off, 'x') || 0) : null
  let y = off ? Number(getAttribute(off, 'y') || 0) : null
  let width = ext ? Number(getAttribute(ext, 'cx') || 0) : null
  let height = ext ? Number(getAttribute(ext, 'cy') || 0) : null

  // A placeholder with no transform of its own takes the layout's, resolved
  // through the placeholder type and index rather than through document order.
  let geometryInherited = false
  if ((x === null || width === null) && placeholder && definition) {
    if (x === null) x = definition.x
    if (y === null) y = definition.y
    if (width === null) width = definition.width
    if (height === null) height = definition.height
    geometryInherited = x !== null && width !== null
  }
  // Neither the slide nor the layout pinned a box: several of R7's own layouts
  // (the title slide among them) rely on the master's placeholder position, so
  // fall back to the conventional region for the placeholder's role. The box is
  // approximate by definition and is reported as inherited.
  if ((x === null || width === null) && placeholder) {
    const fallback = placeholderFallbackBox(placeholder, size)
    if (x === null) x = fallback.x
    if (y === null) y = fallback.y
    if (width === null) width = fallback.width
    if (height === null) height = fallback.height
    geometryInherited = true
  }
  if (x !== null && width !== null && !onSlide) geometryInherited = true

  const rotation = xfrm ? rotationToDegrees(getAttribute(openingTag(xfrm), 'rot')) : 0
  const preset = spPr ? presetOf(spPr) : null

  const type = classify(element.tag, xml, placeholder, preset)
  const textBody = firstElement(xml, 'p:txBody')
  const textModel = textBody
    ? describeTextBody(textBody, context)
    : { text: '', paragraphs: [], alignment: { horizontal: null, vertical: null } }

  const fill = describeFill(spPr)
  const stroke = describeStroke(spPr)
  const image = describeImage(xml, zip, relTargets)

  return {
    id,
    name,
    description,
    type,
    shapeType: type === 'shape' || type === 'textBox' || type === 'placeholder' || type === 'connector'
      ? (preset || null)
      : null,
    preset: preset || null,
    placeholder: placeholder
      ? { type: placeholder.type, idx: placeholder.idx, inherited: geometryInherited }
      : null,
    inherited: inherited || geometryInherited,
    onSlide,
    text: textModel.text,
    x, y, width, height,
    positionValid: x !== null && y !== null && width !== null && height !== null,
    rotation,
    zOrder: position,
    font: textModel.font,
    fill,
    stroke,
    alignment: textModel.alignment,
    paragraphs: textModel.paragraphs,
    image,
    _element: xml,
    _shapeTreeIndex: element.index
  }
}

/** Classify a shape element into the vocabulary a caller reasons about. */
function classify(tag, xml, placeholder, preset) {
  if (tag === 'p:pic') return 'image'
  if (tag === 'p:graphicFrame') {
    if (/<a:tbl>/.test(xml)) return 'table'
    if (/graphicData[^>]*uri="[^"]*chart/.test(xml)) return 'chart'
    if (/graphicData[^>]*uri="[^"]*ole/.test(xml)) return 'embeddedObject'
    return 'graphicFrame'
  }
  if (tag === 'p:grpSp') return 'group'
  if (tag === 'p:contentPart') return 'contentPart'
  if (tag === 'p:cxnSp') return 'connector'

  // A free shape that draws an outline with no fill and no text is the line a
  // user drew. A rectangle with text but no paint is the text box they typed
  // into — the difference is invisible in the markup (both are `<p:sp>` with
  // `prstGeom prst="rect"`), so it is read from what the object actually
  // carries.
  const hasTextBody = Boolean(firstElement(xml, 'p:txBody'))
  if (!placeholder && isOutlineShape(xml, preset, hasTextBody)) return 'connector'
  if (placeholder) return 'placeholder'
  if (hasTextBody && isTextOnlyBox(xml, preset)) return 'textBox'
  return 'shape'
}

/** True for an unpainted rectangle carrying text: what a text box is. */
function isTextOnlyBox(xml, preset) {
  if (preset !== 'rect') return false
  if (!extractElements(xml, 'a:t').some((t) => innerText(t.outerXml).trim() !== '')) return false
  const spPr = firstElement(xml, 'p:spPr')
  if (!spPr) return false
  // Explicit solid/gradient/picture paint means the author drew a shape.
  if (firstElement(spPr, 'a:solidFill') || firstElement(spPr, 'a:gradFill')
    || firstElement(spPr, 'a:blipFill') || firstElement(spPr, 'a:pattFill')) {
    return false
  }
  const line = firstElement(spPr, 'a:ln')
  if (line && !firstElement(line, 'a:noFill')) return false
  return true
}

/**
 * True when an `<p:sp>` is really a drawn line: a line-like preset, no fill
 * and no text of its own.
 */
function isOutlineShape(xml, preset, hasTextBody) {
  const lineLike = preset === 'line' || /^straightConnector/.test(preset || '')
  if (!lineLike) return false
  if (hasTextBody && extractElements(xml, 'a:t').some((t) => innerText(t.outerXml).trim() !== '')) {
    return false
  }
  return true
}

/** The preset geometry name of a shape, if it has one. */
function presetOf(spPr) {
  const prstGeom = firstElement(spPr, 'a:prstGeom')
  if (prstGeom) return getAttribute(openingTag(prstGeom), 'prst') || null
  if (firstElement(spPr, 'a:custGeom')) return 'custom'
  return null
}

/**
 * The conventional box for a placeholder whose position no layout declares.
 *
 * A slide is 16:9 or 4:3 and the regions are proportional, so a fraction of the
 * canvas is stable across both. This is what R7's own layouts produce, and it
 * means "read the geometry" never answers `null` for a title.
 *
 * @param {{type: string}} placeholder
 * @param {{width: number, height: number}} size
 * @returns {{x: number, y: number, width: number, height: number}}
 */
function placeholderFallbackBox(placeholder, size) {
  const width = size && size.width ? size.width : 12192000
  const height = size && size.height ? size.height : 6858000
  const centred = (widthFraction, heightFraction, xFraction, yFraction) => ({
    x: Math.round(width * xFraction),
    y: Math.round(height * yFraction),
    width: Math.round(width * widthFraction),
    height: Math.round(height * heightFraction)
  })

  switch (placeholder.type) {
    case 'title':
      return centred(0.833, 0.185, 0.0835, 0.112)
    case 'ctrTitle':
      return centred(0.75, 0.35, 0.125, 0.16)
    case 'subTitle':
      return centred(0.75, 0.24, 0.125, 0.525)
    case 'body':
      return centred(0.833, 0.55, 0.0835, 0.30)
    case 'pic':
      return centred(0.45, 0.55, 0.52, 0.22)
    case 'tbl':
      return centred(0.833, 0.55, 0.0835, 0.30)
    case 'dt':
      return centred(0.25, 0.05, 0.0417, 0.90)
    case 'ftr':
      return centred(0.5, 0.05, 0.25, 0.90)
    case 'sldNum':
      return centred(0.1, 0.05, 0.86, 0.90)
    default:
      return centred(0.833, 0.5, 0.0835, 0.25)
  }
}

/** Describe the text body: paragraphs, runs, alignment and the effective font. */
function describeTextBody(bodyXml, context) {
  const bodyPr = openingTag(firstElement(bodyXml, 'a:bodyPr') || '')
  const paragraphs = []
  const font = {
    family: null, size: null, bold: null, italic: null, underline: null,
    strike: null, color: null, colorToken: null, transparency: 0
  }

  for (const p of extractElements(bodyXml, 'a:p')) {
    const pPr = firstElement(p.outerXml, 'a:pPr')
    const bullet = bulletOf(pPr)
    const level = pPr ? Number(getAttribute(openingTag(pPr), 'lvl') || 0) : 0
    const style = resolveRunStyle(bodyXml, p.outerXml, context)

    const runs = []
    for (const r of extractElements(p.outerXml, 'a:r')) {
      const rPr = firstElement(r.outerXml, 'a:rPr')
      const colour = rPr ? colourOf(rPr) : null
      const t = firstElement(r.outerXml, 'a:t')
      runs.push({
        text: t ? decodeEntities(innerText(t)) : '',
        family: rPr ? typefaceOf(rPr) : null,
        size: rPr ? sizeOf(rPr) : null,
        bold: rPr ? booleanAttr(openingTag(rPr), 'b') : null,
        italic: rPr ? booleanAttr(openingTag(rPr), 'i') : null,
        underline: rPr ? underlineOf(rPr) : null,
        strike: rPr ? booleanAttr(openingTag(rPr), 'strike') : null,
        color: colour ? colour.hex : null,
        colorToken: colour ? colour.token : null,
        transparency: colour ? colour.transparency : 0
      })
    }
    // `<a:br/>` splits a paragraph visually; the text is joined back with a
    // newline so a caller sees the same string it wrote.
    const text = runs.map((r) => r.text).join('')
      + '\n'.repeat(countLineBreaks(p.outerXml))

    paragraphs.push({
      text,
      level,
      alignment: pPr ? alignmentOf(pPr) : null,
      bullet: bullet.kind,
      bulletCharacter: bullet.character,
      bulletType: bullet.type,
      numberingStart: bullet.startAt,
      marginLeft: pPr ? numberOrNull(getAttribute(openingTag(pPr), 'marL')) : null,
      indent: pPr ? numberOrNull(getAttribute(openingTag(pPr), 'indent')) : null,
      lineSpacing: lineSpacingOf(pPr),
      spaceBefore: spaceOf(pPr, 'a:spcBef'),
      spaceAfter: spaceOf(pPr, 'a:spcAft'),
      runs
    })

    if (font.family === null) {
      font.family = style.family
      font.size = style.size
      font.bold = style.bold
      font.italic = style.italic
      font.underline = style.underline
      font.strike = style.strike
      font.color = style.color
      font.colorToken = style.colorToken
      font.transparency = style.transparency
    }
  }

  return {
    text: paragraphs.map((p) => p.text).join('\n'),
    paragraphs,
    font,
    alignment: {
      horizontal: paragraphs.length > 0 ? paragraphs[0].alignment : null,
      vertical: bodyPr ? anchorOf(bodyPr) : null
    }
  }
}

function countLineBreaks(xml) {
  return (xml.match(/<a:br\s*\/>/g) || []).length
}

/** The inner text of an `<a:t>`, undoing XML entities. */
function innerText(tXml) {
  const match = tXml.match(/^<a:t(?:\s[^>]*)?>([\s\S]*)<\/a:t>$/)
  return match ? match[1] : tXml
}

function decodeEntities(text) {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCodePoint(parseInt(code, 16)))
    .replace(/&amp;/g, '&')
}

function typefaceOf(rPr) {
  const latin = firstElement(rPr, 'a:latin')
  return latin ? getAttribute(openingTag(latin), 'typeface') : null
}

function sizeOf(rPr) {
  const value = getAttribute(openingTag(rPr), 'sz')
  return value === null ? null : Number(value) / 100
}

function underlineOf(rPr) {
  const value = getAttribute(openingTag(rPr), 'u')
  if (value === null || value === 'none') return false
  return value
}

function alignmentOf(pPr) {
  return getAttribute(openingTag(pPr), 'algn') || null
}

function anchorOf(bodyPr) {
  return getAttribute(bodyPr, 'anchor') || null
}

function numberOrNull(value) {
  return value === null ? null : Number(value)
}

function lineSpacingOf(pPr) {
  if (!pPr) return null
  const lnSpc = firstElement(pPr, 'a:lnSpc')
  if (!lnSpc) return null
  const pct = firstElement(lnSpc, 'a:spcPct')
  if (pct) return Number(getAttribute(openingTag(pct), 'val')) / 100000
  const pts = firstElement(lnSpc, 'a:spcPts')
  if (pts) return { points: Number(getAttribute(openingTag(pts), 'val')) / 100 }
  return null
}

function spaceOf(pPr, tag) {
  if (!pPr) return null
  const el = firstElement(pPr, tag)
  if (!el) return null
  const pts = firstElement(el, 'a:spcPts')
  if (pts) return Number(getAttribute(openingTag(pts), 'val')) / 100
  return null
}

/** Bullet state of a paragraph, including the inherited-list case. */
function bulletOf(pPr) {
  if (!pPr) return { kind: 'inherit', character: null, type: null, startAt: null }
  const inner = pPr
  if (firstElement(inner, 'a:buNone')) return { kind: 'none', character: null, type: null, startAt: null }
  const auto = firstElement(inner, 'a:buAutoNum')
  if (auto) {
    return {
      kind: 'numbered',
      character: null,
      type: getAttribute(openingTag(auto), 'type') || 'arabicPeriod',
      startAt: numberOrNull(getAttribute(openingTag(auto), 'startAt'))
    }
  }
  const char = firstElement(inner, 'a:buChar')
  if (char) {
    return {
      kind: 'bullet',
      character: getAttribute(openingTag(char), 'char') || '•',
      type: null,
      startAt: null
    }
  }
  return { kind: 'inherit', character: null, type: null, startAt: null }
}

/** Describe a shape's fill from `<p:spPr>`. */
function describeFill(spPr) {
  if (!spPr) return { color: null, colorToken: null, transparency: 0, kind: null, inherited: true }
  // Direct children only: a shape's `<a:ln>` also carries `<a:noFill>` and
  // `<a:solidFill>`, and a descendant search would report the outline's paint
  // as the shape's fill.
  if (childElement(spPr, 'a:noFill')) {
    return { color: null, colorToken: null, transparency: 0, kind: 'none', inherited: false }
  }
  const solid = childElement(spPr, 'a:solidFill')
  if (solid) {
    const info = colourOf(solid)
    if (info) {
      return {
        color: info.hex,
        colorToken: info.token,
        transparency: info.transparency,
        kind: 'solid',
        inherited: false
      }
    }
  }
  if (childElement(spPr, 'a:gradFill')) {
    return { color: null, colorToken: null, transparency: 0, kind: 'gradient', inherited: false }
  }
  if (childElement(spPr, 'a:blipFill')) {
    return { color: null, colorToken: null, transparency: 0, kind: 'picture', inherited: false }
  }
  if (childElement(spPr, 'a:pattFill')) {
    return { color: null, colorToken: null, transparency: 0, kind: 'pattern', inherited: false }
  }
  return { color: null, colorToken: null, transparency: 0, kind: null, inherited: true }
}

/** Describe a shape's stroke from `<p:spPr>`. */
function describeStroke(spPr) {
  if (!spPr) return { color: null, colorToken: null, width: null, style: null, transparency: 0, none: false }
  const ln = childElement(spPr, 'a:ln')
  if (!ln) return { color: null, colorToken: null, width: null, style: null, transparency: 0, none: false }
  const width = getAttribute(openingTag(ln), 'w')
  if (childElement(ln, 'a:noFill')) {
    return { color: null, colorToken: null, width: width === null ? null : Number(width), style: 'none', transparency: 0, none: true }
  }
  const solid = childElement(ln, 'a:solidFill')
  const info = solid ? colourOf(solid) : null
  const dash = childElement(ln, 'a:prstDash')
  return {
    color: info ? info.hex : null,
    colorToken: info ? info.token : null,
    width: width === null ? null : Number(width),
    style: dash ? (getAttribute(openingTag(dash), 'val') || 'solid') : (info ? 'solid' : null),
    transparency: info ? info.transparency : 0,
    none: false
  }
}

/** Describe an image: relationship, media part, format and rotation lock. */
function describeImage(xml, zip, relTargets) {
  const blip = firstElement(xml, 'a:blip')
  if (!blip) return null
  const relId = getAttribute(blip, 'r:embed') || getAttribute(blip, 'relationships:embed')
  if (!relId) return null
  const rel = relTargets.get(relId)
  const mediaPath = rel ? rel.partPath : null
  const info = mediaPath ? imageInfo(zip, mediaPath) : null
  const picture = firstElement(xml, 'p:cNvPicPr')
  return {
    relId,
    mediaPath,
    format: info ? info.format : null,
    bytes: info ? info.bytes : null,
    hasHyperlink: Boolean(firstElement(blip, 'a:hlinkClick')),
    aspectLocked: picture ? /noChangeAspect="1"/.test(picture) : false
  }
}
