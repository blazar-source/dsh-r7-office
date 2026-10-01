/**
 * pdf-inspect.js — dependency-free PDF text-fidelity inspector (Node stdlib only).
 *
 * Why this exists
 * ---------------
 * R7-Office's `x2t` converter can emit a PDF whose glyphs are *drawn* but whose
 * text is unrecoverable: when the DOCX/PPTX/HTML renderer cannot resolve its font
 * list it emits a fill operation for every glyph with an EMPTY path, so the page
 * is blank and no text object is ever written. A PDF that looks "mostly fine" in a
 * byte dump can therefore be completely text-less. This module measures that
 * condition instead of assuming a conversion worked.
 *
 * Two independent signals are reported:
 *   - `textGlyphs`        glyphs emitted by text-showing operators (`Tj`, `TJ`, `'`, `"`)
 *   - `emptyFillOperators` paint operators (`f`, `B`, `S`, ...) with no path built
 *                          first — the exact signature of glyph outlines that were
 *                          dropped instead of drawn
 *
 * `verdict` is a convenience summary; the raw counters are the authority.
 *
 * Parsing note: object streams are walked SEQUENTIALLY and every `/Length` is
 * honoured, including the indirect form (`/Length 12 0 R`). The indirect form must
 * be tested BEFORE the direct form: a naive `/Length\s+(\d+)/` backtracks on
 * `/Length 12 0 R` and happily reports a length of 1, which silently truncates
 * every compressed stream in the file and hides most objects.
 */

import fs from 'node:fs'
import zlib from 'node:zlib'
import { fontIdentity, readCmap, readGlyph, glyphOutlineKey, sameGlyphSpace } from './sfnt.js'

const WS = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20])
const PATH_CONSTRUCTION_OPS = new Set(['m', 'l', 'c', 'v', 'y', 're', 'h'])
const PATH_PAINT_OPS = new Set(['S', 's', 'f', 'F', 'f*', 'B', 'B*', 'b', 'b*', 'n'])
const PATH_CLIP_OPS = new Set(['W', 'W*'])
const TEXT_SHOW_OPS = new Set(['Tj', 'TJ', "'", '"'])

/** Cyrillic block plus the numero sign, which Russian documents use constantly. */
export function isCyrillicCodePoint(cp) {
  return (cp >= 0x0400 && cp <= 0x04ff) || cp === 0x2116
}

function skipWhitespace(buf, i) {
  while (i < buf.length && WS.has(buf[i])) i++
  return i
}

/**
 * Sequentially scan `N G obj ... endobj` honouring /Length (direct or indirect).
 * Returns objects in file order; every object carries its verbatim dictionary text.
 */
export function parseObjects(buf) {
  const latin = buf.toString('latin1')
  const objects = new Map()
  const order = []
  const numericValues = new Map() // object number -> integer, for indirect /Length
  const re = /(\d{1,10})\s+(\d{1,5})\s+obj\b/g
  let pos = 0
  let headerEnd = 0

  while (pos < buf.length) {
    re.lastIndex = pos
    const m = re.exec(latin)
    if (!m) break
    if (order.length === 0) headerEnd = m.index
    const num = parseInt(m[1], 10)
    const gen = parseInt(m[2], 10)
    const dictStart = m.index + m[0].length

    const streamKeyword = latin.indexOf('stream', dictStart)
    const endObjKeyword = latin.indexOf('endobj', dictStart)
    const hasStream = streamKeyword !== -1 && (endObjKeyword === -1 || streamKeyword < endObjKeyword)

    if (!hasStream) {
      const dictEnd = endObjKeyword === -1 ? buf.length : endObjKeyword
      const dict = latin.slice(dictStart, dictEnd)
      objects.set(num, { num, gen, dict, hasStream: false, raw: null, dataStart: -1, dataEnd: -1 })
      order.push(num)
      const trimmed = dict.trim()
      if (/^-?\d+$/.test(trimmed)) numericValues.set(num, parseInt(trimmed, 10))
      pos = dictEnd + 6
      continue
    }

    const dict = latin.slice(dictStart, streamKeyword)
    let dataStart = streamKeyword + 6
    if (buf[dataStart] === 0x0d) dataStart++
    if (buf[dataStart] === 0x0a) dataStart++

    let length = null
    const indirect = /\/Length\s+(\d+)\s+\d+\s+R/.exec(dict)
    if (indirect) {
      const ref = parseInt(indirect[1], 10)
      if (numericValues.has(ref)) length = numericValues.get(ref)
    } else {
      const direct = /\/Length\s+(\d+)/.exec(dict)
      if (direct) length = parseInt(direct[1], 10)
    }

    let dataEnd
    if (length !== null && dataStart + length <= buf.length) {
      dataEnd = dataStart + length
    } else {
      // Forward /Length reference (x2t writes the length object AFTER its stream).
      // Accept only an `endstream` that is immediately followed by `endobj`, so a
      // literal "endstream" inside compressed bytes cannot split the object. The
      // end-of-line marker before the keyword belongs to the file, not the data.
      let probe = dataStart
      dataEnd = buf.length
      for (;;) {
        const e = latin.indexOf('endstream', probe)
        if (e === -1) break
        const after = skipWhitespace(buf, e + 9)
        if (latin.startsWith('endobj', after)) {
          dataEnd = e
          if (dataEnd > dataStart && buf[dataEnd - 1] === 0x0a) dataEnd--
          if (dataEnd > dataStart && buf[dataEnd - 1] === 0x0d) dataEnd--
          break
        }
        probe = e + 9
      }
    }

    objects.set(num, {
      num,
      gen,
      dict,
      hasStream: true,
      raw: buf.slice(dataStart, Math.min(dataEnd, buf.length)),
      dataStart,
      dataEnd
    })
    order.push(num)
    const next = latin.indexOf('endobj', Math.max(dataEnd, dataStart))
    pos = next === -1 ? Math.max(dataEnd, dataStart + 1) : next + 6
  }

  return {
    header: latin.slice(0, headerEnd),
    objects,
    order,
    numericValues,
    trailerText: extractTrailerText(latin)
  }
}

function extractTrailerText(latin) {
  const idx = latin.lastIndexOf('trailer')
  if (idx === -1) return null
  const start = latin.indexOf('<<', idx)
  if (start === -1) return null
  let depth = 0
  for (let i = start; i < latin.length - 1; i++) {
    if (latin[i] === '<' && latin[i + 1] === '<') { depth++; i++; continue }
    if (latin[i] === '>' && latin[i + 1] === '>') { depth--; i++; if (depth === 0) return latin.slice(start, i + 1); continue }
  }
  return null
}

/** Decode one stream object (FlateDecode / ASCIIHexDecode). Returns null when undecodable. */
export function decodeStream(obj) {
  if (!obj || !obj.hasStream || !obj.raw) return null
  if (obj.decoded !== undefined) return obj.decoded
  const filters = /\/Filter\s*(\[[^\]]*\]|\/\w+)/.exec(obj.dict)
  const filterText = filters ? filters[1] : ''
  let out = obj.raw
  try {
    if (/FlateDecode/.test(filterText)) out = zlib.inflateSync(out)
    else if (/ASCIIHexDecode/.test(filterText)) {
      const hex = out.toString('latin1').replace(/\s/g, '').replace(/>.*$/, '')
      out = Buffer.from(hex.length % 2 ? hex + '0' : hex, 'hex')
    }
  } catch {
    out = null
  }
  obj.decoded = out
  return out
}

function decodeUtf16BEMulti(hex) {
  let out = ''
  for (let i = 0; i + 4 <= hex.length; i += 4) {
    out += String.fromCharCode(parseInt(hex.slice(i, i + 4), 16))
  }
  return out
}

/**
 * Parse the mappings of a ToUnicode CMap (beginbfchar / beginbfrange, including the
 * array form of bfrange).
 * @returns {Array<{code: number, unicode: string}>}
 */
export function parseToUnicodeCMap(text) {
  const out = []
  const bfcharRe = /(\d+)\s+beginbfchar([\s\S]*?)endbfchar/g
  let m
  while ((m = bfcharRe.exec(text))) {
    const pairRe = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g
    let p
    while ((p = pairRe.exec(m[2]))) {
      out.push({ code: parseInt(p[1], 16), unicode: decodeUtf16BEMulti(p[2]) })
    }
  }
  const bfrangeRe = /(\d+)\s+beginbfrange([\s\S]*?)endbfrange/g
  while ((m = bfrangeRe.exec(text))) {
    for (const line of m[2].split(/\r?\n/)) {
      const single = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(line)
      if (single) {
        const lo = parseInt(single[1], 16)
        const hi = parseInt(single[2], 16)
        const dst = parseInt(single[3], 16)
        for (let c = lo; c <= hi && c - lo < 65536; c++) {
          out.push({ code: c, unicode: String.fromCodePoint(dst + (c - lo)) })
        }
        continue
      }
      const array = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([\s\S]*?)\]/.exec(line)
      if (array) {
        const lo = parseInt(array[1], 16)
        const items = array[3].match(/<([0-9A-Fa-f]+)>/g) || []
        items.forEach((item, index) => {
          out.push({ code: lo + index, unicode: decodeUtf16BEMulti(item.replace(/[<>]/g, '')) })
        })
      }
    }
  }
  return out
}

/**
 * Count what a page content stream actually draws.
 *
 * `emptyFillOperators` counts paint operators that had no path to paint — the
 * signature of text lost between layout and PDF emission.
 */
export function analyzeContentStream(text) {
  let textGlyphs = 0
  let textRuns = 0
  let textShowOperators = 0
  let pathOperators = 0
  let emptyFillOperators = 0
  let pathConstructed = false
  let inText = false
  const operands = []
  const n = text.length
  let i = 0

  const opRe = /[A-Za-z*'"][A-Za-z0-9*'"]*/g

  while (i < n) {
    const ch = text[i]

    if (ch === '%') { // comment to end of line
      const eol = text.indexOf('\n', i)
      i = eol === -1 ? n : eol + 1
      continue
    }
    if (ch === '<' && text[i + 1] === '<') { operands.push({ t: 'dict' }); i += 2; continue }
    if (ch === '<') {
      const e = text.indexOf('>', i)
      if (e === -1) break
      operands.push({ t: 'hex', v: text.slice(i + 1, e), glyphs: Math.floor(text.slice(i + 1, e).replace(/\s/g, '').length / 4) })
      i = e + 1
      continue
    }
    if (ch === '(') {
      let depth = 1
      let j = i + 1
      while (j < n && depth > 0) {
        if (text[j] === '\\') j += 2
        else if (text[j] === '(') { depth++; j++ }
        else if (text[j] === ')') { depth--; j++ }
        else j++
      }
      operands.push({ t: 'str', v: text.slice(i + 1, Math.max(i + 1, j - 1)), glyphs: Math.max(0, j - i - 2) })
      i = j
      continue
    }
    if (ch === '[') { operands.push({ t: 'arrOpen' }); i++; continue }
    if (ch === ']') {
      const items = []
      while (operands.length && operands[operands.length - 1].t !== 'arrOpen') items.unshift(operands.pop())
      if (operands.length) operands.pop()
      let glyphs = 0
      for (const item of items) {
        if (item.t === 'hex' || item.t === 'str') glyphs += item.glyphs
        else if (item.t === 'arr') glyphs += item.glyphs
      }
      operands.push({ t: 'arr', glyphs })
      i++
      continue
    }
    if (ch === '/' ) {
      const e = i + 1
      let j = e
      while (j < n && !/[\s/<>[\]()]/.test(text[j])) j++
      operands.push({ t: 'name', v: text.slice(e, j) })
      i = j
      continue
    }
    if (/\s/.test(ch)) { i++; continue }

    opRe.lastIndex = i
    const m = opRe.exec(text)
    if (m && m.index === i) {
      const op = m[0]
      if (op === 'BI') { // inline image: skip to EI
        const ei = text.indexOf('EI', text.indexOf('ID', i) + 2)
        i = ei === -1 ? n : ei + 2
        operands.length = 0
        continue
      }
      if (op === 'BT') inText = true
      else if (op === 'ET') inText = false
      else if (inText && TEXT_SHOW_OPS.has(op)) {
        let glyphs = 0
        for (const item of operands) {
          if (item.t === 'hex' || item.t === 'str' || item.t === 'arr') glyphs += item.glyphs || 0
        }
        textShowOperators++
        if (glyphs > 0) { textGlyphs += glyphs; textRuns++ }
      } else if (!inText && PATH_CONSTRUCTION_OPS.has(op)) {
        pathConstructed = true
        pathOperators++
      } else if (!inText && PATH_CLIP_OPS.has(op)) {
        pathOperators++
      } else if (!inText && PATH_PAINT_OPS.has(op)) {
        pathOperators++
        if (!pathConstructed) emptyFillOperators++
        pathConstructed = false
      }
      operands.length = 0
      i += op.length
      continue
    }

    const num = /^[-+]?(?:\d+\.?\d*|\.\d+)/.exec(text.slice(i, i + 40))
    if (num) { operands.push({ t: 'num' }); i += num[0].length; continue }
    i++
  }

  return { textGlyphs, textRuns, textShowOperators, pathOperators, emptyFillOperators }
}

function resolveRef(objects, dict, key) {
  const re = new RegExp('/' + key + '\\s+(\\d+)\\s+\\d+\\s+R')
  const m = re.exec(dict)
  if (!m) return null
  return objects.get(parseInt(m[1], 10)) || null
}

function resolveRefArray(objects, dict, key) {
  const re = new RegExp('/' + key + '\\s*\\[([^\\]]*)\\]')
  const m = re.exec(dict)
  if (!m) return []
  const out = []
  for (const r of m[1].match(/(\d+)\s+\d+\s+R/g) || []) out.push(objects.get(parseInt(r, 10)) || null)
  return out
}

function refNumber(dict, key) {
  const m = new RegExp('/' + key + '\\s+(\\d+)\\s+\\d+\\s+R').exec(dict)
  return m ? parseInt(m[1], 10) : null
}

function nameValue(dict, key) {
  const m = new RegExp('/' + key + '\\s*/([^\\s/<>\\[\\]]+)').exec(dict)
  return m ? m[1] : null
}

function hasFontFile(objects, descriptor) {
  if (!descriptor) return false
  if (/\/FontFile[23]?\b/.test(descriptor.dict)) return true
  const ref = refNumber(descriptor.dict, 'FontFile') ?? refNumber(descriptor.dict, 'FontFile2') ?? refNumber(descriptor.dict, 'FontFile3')
  return ref !== null && objects.has(ref)
}

// ---------------------------------------------------------------------------
// Glyph-chain verification
//
// A CIDFontType2 PDF font is only *correct* when both of these land on the same
// outline:
//
//   content code -> CID -> CIDToGIDMap[cid] -> GID -> glyf outline
//   content code -> CID -> ToUnicode[cid]  -> Unicode -> cmap -> GID
//
// Text extraction only exercises the second chain, so a PDF whose glyph IDs were
// taken from a DIFFERENT font (R7's metric-compatible substitute) extracts
// perfectly and draws the wrong shapes. Deciding which glyph a GID holds needs
// the font the subset was cut from - the subset's own `cmap` cannot say, because
// x2t writes a private format 6 table whose glyphIdArray is a copy of
// CIDToGIDMap. So the caller supplies candidate source fonts (the same font list
// handed to x2t) and this check stays renderer-free.
// ---------------------------------------------------------------------------

/** Cache parsed candidate fonts; a font list can hold hundreds of files. */
const SOURCE_FONT_CACHE = new Map()

function sourceFontRecord(filePath) {
  if (SOURCE_FONT_CACHE.has(filePath)) return SOURCE_FONT_CACHE.get(filePath)
  let record = null
  try {
    const buf = fs.readFileSync(filePath)
    const identity = fontIdentity(buf)
    if (identity) record = { path: filePath, identity, cmap: readCmap(identity.sfnt) }
  } catch {
    record = null
  }
  SOURCE_FONT_CACHE.set(filePath, record)
  return record
}

/** Clear the parsed-candidate cache (tests and long-lived hosts). */
export function clearSourceFontCache() {
  SOURCE_FONT_CACHE.clear()
}

function fontBasename(filePath) {
  return String(filePath).split(/[\\/]/).pop() || ''
}

/** First run of letters in a family name, used to prefilter candidate files. */
function familyToken(family) {
  const match = /[A-Za-z]{3,}/.exec(family || '')
  return match ? match[0].toLowerCase() : null
}

/**
 * Candidate source fonts that share the embedded subset's glyph space.
 * `sameGlyphSpace` requires the same glyph count and the same advance width for
 * every glyph, so a different version of the same family is rejected rather than
 * trusted (and a rejected candidate yields "unverifiable", never "corrupt").
 */
export function findSourceFonts(identity, fontPaths) {
  const token = familyToken(identity.family)
  const matches = []
  for (const candidate of fontPaths || []) {
    if (!candidate) continue
    if (token && !fontBasename(candidate).toLowerCase().includes(token)) continue
    const record = sourceFontRecord(candidate)
    if (!record) continue
    if (sameGlyphSpace(identity, record.identity)) matches.push(record)
  }
  return matches
}

/**
 * Verify the CID/GID chain of every embedded CIDFontType2 font.
 *
 * @param {ReturnType<typeof parseObjects>} parsed
 * @param {object} [options]
 * @param {string[]} [options.fontPaths] candidate source font files
 * @param {number} [options.maxExamples=5]
 * @returns {{
 *   checked: boolean, consistent: boolean|null, sourceFontsMatched: number,
 *   cidsChecked: number, mismatches: number,
 *   fonts: Array<{object: number, baseFont: string|null, sourceFont: string|null,
 *                 cidsChecked: number, mismatches: number,
 *                 examples: Array<{cid: number, character: string, drawnGid: number, expectedGid: number}>,
 *                 reason?: string}>
 * }}
 */
export function analyzeGlyphChain(parsed, options = {}) {
  const { objects, order } = parsed
  const fontPaths = options.fontPaths || []
  const maxExamples = options.maxExamples ?? 5
  const fonts = []
  let cidsChecked = 0
  let mismatches = 0
  let sourceFontsMatched = 0

  for (const num of order) {
    const obj = objects.get(num)
    if (!/\/Type\s*\/Font\b/.test(obj.dict)) continue
    if (!/\/Subtype\s*\/Type0\b/.test(obj.dict)) continue

    const baseFont = nameValue(obj.dict, 'BaseFont')
    const report = { object: num, baseFont, sourceFont: null, cidsChecked: 0, mismatches: 0, examples: [] }

    const descendant = resolveRefArray(objects, obj.dict, 'DescendantFonts')[0]
    if (!descendant) { report.reason = 'no-descendant-font'; fonts.push(report); continue }
    const descriptorRef = refNumber(descendant.dict, 'FontDescriptor')
    const descriptor = descriptorRef === null ? null : objects.get(descriptorRef)
    const fileRef = descriptor
      ? (refNumber(descriptor.dict, 'FontFile2') ?? refNumber(descriptor.dict, 'FontFile3'))
      : null
    if (fileRef === null) { report.reason = 'no-embedded-font-file'; fonts.push(report); continue }

    const fontFile = decodeStream(objects.get(fileRef))
    const identity = fontFile ? fontIdentity(fontFile) : null
    if (!identity) { report.reason = 'unsupported-font-file'; fonts.push(report); continue }

    const candidates = findSourceFonts(identity, fontPaths)
    if (!candidates.length) { report.reason = 'no-matching-source-font'; fonts.push(report); continue }
    sourceFontsMatched++
    const source = candidates[0]
    report.sourceFont = source.path

    const toUnicode = resolveRef(objects, obj.dict, 'ToUnicode')
    const cidMapObject = descriptor === null ? null : resolveRef(objects, descendant.dict, 'CIDToGIDMap')
    const cidMapStream = cidMapObject ? decodeStream(cidMapObject) : null
    if (!toUnicode || !cidMapStream) { report.reason = 'no-cid-mapping'; fonts.push(report); continue }

    const decodedToUnicode = decodeStream(toUnicode)
    if (!decodedToUnicode) { report.reason = 'no-cid-mapping'; fonts.push(report); continue }
    const cidToGid = []
    for (let i = 0; i + 2 <= cidMapStream.length; i += 2) cidToGid.push(cidMapStream.readUInt16BE(i))

    const outlineCache = new Map()
    const embeddedOutlineKey = (gid) => {
      if (outlineCache.has(gid)) return outlineCache.get(gid)
      const key = glyphOutlineKey(readGlyph(identity.sfnt, gid))
      outlineCache.set(gid, key)
      return key
    }
    const sourceOutlineCache = new Map()
    const sourceOutlineKey = (gid) => {
      if (sourceOutlineCache.has(gid)) return sourceOutlineCache.get(gid)
      const key = glyphOutlineKey(readGlyph(source.identity.sfnt, gid))
      sourceOutlineCache.set(gid, key)
      return key
    }

    for (const pair of parseToUnicodeCMap(decodedToUnicode.toString('latin1'))) {
      const cps = [...pair.unicode]
      if (cps.length !== 1) continue
      const codePoint = cps[0].codePointAt(0)
      if (codePoint === 0) continue
      const expectedGid = source.cmap.get(codePoint)
      if (expectedGid === undefined) continue
      if (pair.code >= cidToGid.length) continue
      const drawnGid = cidToGid[pair.code]
      report.cidsChecked++
      if (drawnGid === expectedGid) continue
      // A different index is fine when it holds the identical outline (many
      // fonts share one glyph between look-alike characters).
      const drawnKey = embeddedOutlineKey(drawnGid)
      const expectedKey = sourceOutlineKey(expectedGid)
      if (drawnKey && expectedKey && drawnKey === expectedKey) continue
      report.mismatches++
      if (report.examples.length < maxExamples) {
        report.examples.push({ cid: pair.code, character: cps[0], drawnGid, expectedGid })
      }
    }

    cidsChecked += report.cidsChecked
    mismatches += report.mismatches
    fonts.push(report)
  }

  return {
    checked: cidsChecked > 0,
    consistent: cidsChecked === 0 ? null : mismatches === 0,
    sourceFontsMatched,
    cidsChecked,
    mismatches,
    fonts
  }
}

/**
 * Verify the CID/GID chain of a PDF file or buffer.
 * @param {string|Buffer} input
 * @param {{fontPaths?: string[], maxExamples?: number}} [options]
 */
export function inspectGlyphChain(input, options = {}) {
  const buf = Buffer.isBuffer(input) ? input : fs.readFileSync(input)
  return analyzeGlyphChain(parseObjects(buf), options)
}

/**
 * Inspect a PDF file (or byte buffer) for text fidelity.
 *
 * @param {string|Buffer} input path to a PDF, or the PDF bytes
 * @param {object} [options]
 * @param {string[]} [options.fontPaths] candidate source fonts (the same font list
 *   x2t was given). When supplied, every embedded CID font's glyph chain is
 *   verified against them and reported in `glyphChain`. Without it the chain is
 *   simply not checkable and `glyphChain.checked` is false.
 * @returns {{
 *   filePath: string|null, bytes: number, objects: number, pages: number,
 *   fonts: Array<{baseFont: string|null, subtype: string|null, encoding: string|null,
 *                 embedded: boolean, descendantSubtype: string|null,
 *                 toUnicodeMappings: number, cyrillicMappings: number,
 *                 glyphChainMismatches: number, glyphChainChecked: number,
 *                 glyphChainSourceFont: string|null}>,
 *   toUnicodeMappings: number, cyrillicMappings: number,
 *   textGlyphs: number, textRuns: number, pathOperators: number, emptyFillOperators: number,
 *   pageStats: Array<{index: number, textGlyphs: number, pathOperators: number, emptyFillOperators: number}>,
 *   hasExtractableText: boolean, hasCyrillicText: boolean, textLossSuspected: boolean,
 *   verdict: 'text'|'outlined'|'mixed', malformedToUnicode: Array<{object: number, declared: number, actual: number}>,
 *   glyphChain: ReturnType<typeof analyzeGlyphChain>, glyphChainChecked: boolean,
 *   glyphChainMismatches: number, glyphMapInconsistent: boolean
 * }}
 */
export function inspectPdf(input, options = {}) {
  const isBuffer = Buffer.isBuffer(input)
  const buf = isBuffer ? input : fs.readFileSync(input)
  const parsed = parseObjects(buf)
  const { objects, order } = parsed

  // ---- fonts (one entry per logical font; descendants folded into their parent) ----
  const descendantNumbers = new Set()
  for (const num of order) {
    const obj = objects.get(num)
    if (!/\/Type\s*\/Font\b/.test(obj.dict)) continue
    for (const d of resolveRefArray(objects, obj.dict, 'DescendantFonts')) {
      if (d) descendantNumbers.add(d.num)
    }
  }

  const fonts = []
  let toUnicodeMappings = 0
  let cyrillicMappings = 0
  const malformedToUnicode = []

  for (const num of order) {
    const obj = objects.get(num)
    if (!/\/Type\s*\/Font\b/.test(obj.dict)) continue
    if (descendantNumbers.has(num)) continue

    const baseFont = nameValue(obj.dict, 'BaseFont')
    const subtype = nameValue(obj.dict, 'Subtype')
    const encoding = nameValue(obj.dict, 'Encoding')

    let descriptor = resolveRef(objects, obj.dict, 'FontDescriptor')
    let descendantSubtype = null
    const descendants = resolveRefArray(objects, obj.dict, 'DescendantFonts')
    if (descendants.length && descendants[0]) {
      descendantSubtype = nameValue(descendants[0].dict, 'Subtype')
      descriptor = resolveRef(objects, descendants[0].dict, 'FontDescriptor') || descriptor
    }

    let mappings = 0
    let cyrillic = 0
    const toUnicode = resolveRef(objects, obj.dict, 'ToUnicode')
    if (toUnicode) {
      const decoded = decodeStream(toUnicode)
      if (decoded) {
        const text = decoded.toString('latin1')
        const pairs = parseToUnicodeCMap(text)
        mappings = pairs.length
        for (const pair of pairs) {
          if ([...pair.unicode].some(ch => isCyrillicCodePoint(ch.codePointAt(0)))) cyrillic++
        }
        for (const block of findMalformedCMapBlocks(text)) malformedToUnicode.push({ object: toUnicode.num, ...block })
      }
    }

    toUnicodeMappings += mappings
    cyrillicMappings += cyrillic
    fonts.push({
      object: num,
      baseFont,
      subtype,
      encoding,
      embedded: hasFontFile(objects, descriptor),
      descendantSubtype,
      toUnicodeMappings: mappings,
      cyrillicMappings: cyrillic
    })
  }

  // ---- pages ----
  const pageStats = []
  let textGlyphs = 0
  let textRuns = 0
  let pathOperators = 0
  let emptyFillOperators = 0

  for (const num of order) {
    const obj = objects.get(num)
    if (!/\/Type\s*\/Page\b/.test(obj.dict)) continue
    if (/\/Type\s*\/Pages\b/.test(obj.dict)) continue

    const contents = []
    const arrayRef = /\/Contents\s*\[([^\]]*)\]/.exec(obj.dict)
    if (arrayRef) {
      for (const r of arrayRef[1].match(/(\d+)\s+\d+\s+R/g) || []) contents.push(objects.get(parseInt(r, 10)))
    } else {
      contents.push(resolveRef(objects, obj.dict, 'Contents'))
    }

    const stats = { index: pageStats.length + 1, textGlyphs: 0, pathOperators: 0, emptyFillOperators: 0 }
    for (const stream of contents) {
      if (!stream) continue
      const decoded = decodeStream(stream)
      if (!decoded) continue
      const analysis = analyzeContentStream(decoded.toString('latin1'))
      stats.textGlyphs += analysis.textGlyphs
      stats.pathOperators += analysis.pathOperators
      stats.emptyFillOperators += analysis.emptyFillOperators
      textGlyphs += analysis.textGlyphs
      textRuns += analysis.textRuns
      pathOperators += analysis.pathOperators
      emptyFillOperators += analysis.emptyFillOperators
    }
    pageStats.push(stats)
  }

  // A page that declares glyph fills but paints nothing is text that was dropped:
  // require a meaningful number of empty paints so an image-only page is not flagged.
  const textLossSuspected = emptyFillOperators >= 20 && textGlyphs <= emptyFillOperators / 4

  // ---- glyph chain (only when the caller can name the source fonts) ----
  const glyphChain = analyzeGlyphChain(parsed, { fontPaths: options.fontPaths || [] })
  const chainByObject = new Map(glyphChain.fonts.map((entry) => [entry.object, entry]))
  for (const font of fonts) {
    const chain = chainByObject.get(font.object)
    font.glyphChainChecked = chain ? chain.cidsChecked : 0
    font.glyphChainMismatches = chain ? chain.mismatches : 0
    font.glyphChainSourceFont = chain ? chain.sourceFont : null
  }

  let verdict
  if (pageStats.length === 0) verdict = textGlyphs > 0 ? 'text' : 'outlined'
  else if (textGlyphs === 0) verdict = 'outlined'
  else if (textLossSuspected) verdict = 'mixed'
  else verdict = 'text'

  return {
    filePath: isBuffer ? null : input,
    bytes: buf.length,
    objects: objects.size,
    pages: pageStats.length,
    fonts,
    toUnicodeMappings,
    cyrillicMappings,
    textGlyphs,
    textRuns,
    pathOperators,
    emptyFillOperators,
    pageStats,
    hasExtractableText: textGlyphs > 0,
    hasCyrillicText: cyrillicMappings > 0,
    textLossSuspected,
    verdict,
    malformedToUnicode,
    glyphChain,
    glyphChainChecked: glyphChain.checked,
    glyphChainMismatches: glyphChain.mismatches,
    glyphMapInconsistent: glyphChain.mismatches > 0
  }
}

/** Report `beginbfchar`/`beginbfrange` blocks whose declared entry count is wrong. */
export function findMalformedCMapBlocks(text) {
  const problems = []
  const blocks = [
    { keyword: 'beginbfchar', endKeyword: 'endbfchar' },
    { keyword: 'beginbfrange', endKeyword: 'endbfrange' }
  ]
  for (const { keyword, endKeyword } of blocks) {
    const re = new RegExp('(\\d+)\\s+' + keyword + '([\\s\\S]*?)' + endKeyword, 'g')
    let m
    while ((m = re.exec(text))) {
      const declared = parseInt(m[1], 10)
      const actual = countCMapEntries(keyword, m[2])
      if (declared !== actual) problems.push({ keyword, declared, actual })
    }
  }
  return problems
}

function countCMapEntries(keyword, body) {
  if (keyword === 'beginbfchar') {
    return (body.match(/<[0-9A-Fa-f]+>\s*<[0-9A-Fa-f]+>/g) || []).length
  }
  let count = 0
  for (const line of body.split(/\r?\n/)) {
    const single = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(line)
    if (single) {
      count += parseInt(single[2], 16) - parseInt(single[1], 16) + 1
      continue
    }
    const array = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([\s\S]*?)\]/.exec(line)
    if (array) count += (array[3].match(/<[0-9A-Fa-f]+>/g) || []).length
  }
  return count
}

/**
 * Repair malformed ToUnicode CMaps in place.
 *
 * R7's PDF writer sometimes declares the wrong `beginbfchar` count (observed:
 * "944 beginbfchar" wrapping 94 entries) for the font that carries the body text.
 * Conforming readers reject the block, so the text renders but copies out as raw
 * CIDs. Rewriting the count is a structural repair of a provably wrong number; it
 * never invents text.
 *
 * @param {string} filePath
 * @returns {{repaired: boolean, repairs: Array<{object:number,keyword:string,declared:number,actual:number}>}}
 */
export function repairToUnicodeCMaps(filePath) {
  const buf = fs.readFileSync(filePath)
  const parsed = parseObjects(buf)
  const { objects, order } = parsed
  const repairs = []
  const patched = new Map()

  const cMapObjects = new Set()
  for (const num of order) {
    const obj = objects.get(num)
    if (!/\/Type\s*\/Font\b/.test(obj.dict)) continue
    const ref = refNumber(obj.dict, 'ToUnicode')
    if (ref !== null) cMapObjects.add(ref)
  }

  for (const num of cMapObjects) {
    const obj = objects.get(num)
    const decoded = decodeStream(obj)
    if (!decoded) continue
    let text = decoded.toString('latin1')
    const problems = findMalformedCMapBlocks(text)
    if (!problems.length) continue
    let changed = false
    for (const problem of problems) {
      const re = new RegExp('\\d+\\s+' + problem.keyword)
      const fixed = text.replace(re, problem.actual + ' ' + problem.keyword)
      if (fixed !== text) {
        text = fixed
        changed = true
        repairs.push({ object: num, keyword: problem.keyword, declared: problem.declared, actual: problem.actual })
      }
    }
    if (changed) patched.set(num, Buffer.from(text, 'latin1'))
  }

  if (!patched.size) return { repaired: false, repairs: [] }

  const rebuilt = rebuildPdf(parsed, patched)
  const tmp = filePath + '.pdf-repair-tmp'
  fs.writeFileSync(tmp, rebuilt)
  fs.renameSync(tmp, filePath)
  return { repaired: true, repairs }
}

/**
 * Re-serialise a parsed PDF with replacement (uncompressed) stream payloads and a
 * fresh classic cross-reference table. Untouched objects keep their original bytes
 * and dictionaries, so /Length references between them stay valid.
 */
export function rebuildPdf(parsed, replacedStreams = new Map()) {
  const { objects, order, header, trailerText } = parsed
  const maxNum = Math.max(...order, 0)
  const size = maxNum + 1

  const rootRef = refNumber(trailerText || '', 'Root') ?? findRefInFile(objects, order, 'Root') ?? findCatalogNumber(objects, order)
  const infoRef = refNumber(trailerText || '', 'Info') ?? findRefInFile(objects, order, 'Info')
  const idText = trailerText ? (/\/ID\s*\[[^\]]*\]/.exec(trailerText) || [])[0] : undefined

  const chunks = []
  let offset = 0
  const push = (value) => {
    const b = Buffer.isBuffer(value) ? value : Buffer.from(value, 'latin1')
    chunks.push(b)
    offset += b.length
  }

  push(header.endsWith('\n') ? header : header + '\n')

  const offsets = new Map()
  const sorted = [...order].sort((a, b) => a - b)
  for (const num of sorted) {
    const obj = objects.get(num)
    offsets.set(num, offset)
    let dict = obj.dict.replace(/\s*$/, '')

    if (replacedStreams.has(num)) {
      const data = zlib.deflateSync(replacedStreams.get(num))
      const withoutLength = dict.replace(/\/Length\s+\d+(\s+\d+\s+R)?/, '')
      const open = withoutLength.indexOf('<<')
      // /Length must live INSIDE the stream dictionary; a stray token before `<<`
      // makes readers skip the object entirely.
      dict = open === -1
        ? `/Length ${data.length}\n${withoutLength}`
        : withoutLength.slice(0, open + 2) + `\n/Length ${data.length}` + withoutLength.slice(open + 2)
      push(`${num} ${obj.gen} obj\n${dict}\nstream\n`)
      push(data)
      push('\nendstream\nendobj\n')
      continue
    }

    if (obj.hasStream) {
      push(`${num} ${obj.gen} obj\n${dict}\nstream\n`)
      push(obj.raw)
      push('\nendstream\nendobj\n')
      continue
    }
    push(`${num} ${obj.gen} obj\n${dict}\nendobj\n`)
  }

  const xrefOffset = offset
  const lines = ['xref', `0 ${size}`]
  for (let i = 0; i < size; i++) {
    if (offsets.has(i)) lines.push(String(offsets.get(i)).padStart(10, '0') + ' 00000 n ')
    else lines.push('0000000000 65535 f ')
  }
  const trailerParts = [`/Size ${size}`]
  if (rootRef !== null && rootRef !== undefined) trailerParts.push(`/Root ${rootRef} 0 R`)
  if (infoRef !== null && infoRef !== undefined) trailerParts.push(`/Info ${infoRef} 0 R`)
  if (idText) trailerParts.push(idText)
  lines.push('trailer', `<< ${trailerParts.join(' ')} >>`, 'startxref', String(xrefOffset), '%%EOF', '')
  push(lines.join('\n'))

  return Buffer.concat(chunks)
}

function findRefInFile(objects, order, key) {
  for (const num of order) {
    const m = new RegExp('/' + key + '\\s+(\\d+)\\s+\\d+\\s+R').exec(objects.get(num).dict)
    if (m) return parseInt(m[1], 10)
  }
  return null
}

function findCatalogNumber(objects, order) {
  for (const num of order) {
    if (/\/Type\s*\/Catalog\b/.test(objects.get(num).dict)) return num
  }
  return null
}
