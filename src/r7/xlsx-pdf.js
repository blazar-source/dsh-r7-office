/**
 * PDF plumbing for the XLSX → PDF path.
 *
 * Why this exists: R7's `x2t` converter writes the PDF for exactly ONE
 * worksheet per run — whichever tab `xl/workbook.xml` marks active — and its
 * CLI has no "all sheets" switch (the params-XML form exposes file/font/format
 * fields only, and neither multiple `tabSelected` tabs nor per-sheet print
 * areas change the outcome; measured, see the module's tests). One workbook is
 * therefore rendered once per worksheet with a scratch copy that activates that
 * tab, and the resulting single-sheet PDFs are concatenated here, in tab order.
 *
 * The concatenation is deliberately minimal and structural: every indirect
 * object of every input is copied verbatim (bytes included — content streams,
 * embedded fonts and ToUnicode maps are never decoded or rewritten) with only
 * its object number changed, and a fresh page tree, catalog, xref table and
 * trailer are written around the copies. Because nothing inside an object is
 * reinterpreted, what a reader could extract from a single-sheet PDF it can
 * still extract from the merged one.
 *
 * The inputs this handles are the ones `x2t` produces: a classic object layout
 * (no object streams) with a cross-reference stream. Objects are found by
 * scanning for `N G obj`, which does not depend on the xref variant; a file
 * with compressed object streams (`/ObjStm`) is rejected rather than silently
 * mangled.
 *
 * This file is XLSX-subsystem plumbing, not a general-purpose PDF library: it
 * merges, counts pages, and does nothing else.
 */

import { extractElements } from '../shared/xml.js'

/** Token delimiters of the PDF grammar, used to find the end of a name. */
const NAME_END = /[\s()<>[\]{}/%]/

/** Whether a character can follow a complete keyword token. */
function isDelimiter(ch) {
  return ch === undefined || NAME_END.test(ch)
}

/** Skip whitespace and `%` comments. */
function skipWhitespace(text, pos) {
  let i = pos
  while (i < text.length) {
    const ch = text[i]
    if (ch === '%') {
      while (i < text.length && text[i] !== '\n' && text[i] !== '\r') i++
      continue
    }
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t' || ch === '\f' || ch === '\0') {
      i++
      continue
    }
    break
  }
  return i
}

/** Read a `/Name` token. `#xx` escapes are carried through untouched. */
function readName(text, pos) {
  let i = pos + 1
  while (i < text.length && !NAME_END.test(text[i])) i++
  return { value: { kind: 'token', text: text.slice(pos, i) }, end: i }
}

/** Read a `(literal string)`, honouring nested parentheses and escapes. */
function readLiteralString(text, pos) {
  let i = pos + 1
  let depth = 1
  while (i < text.length) {
    const ch = text[i]
    if (ch === '\\') {
      i += 2
      continue
    }
    if (ch === '(') depth++
    else if (ch === ')') {
      depth--
      if (depth === 0) {
        i++
        break
      }
    }
    i++
  }
  return { value: { kind: 'token', text: text.slice(pos, i) }, end: i }
}

/** Read a `<hex string>`. */
function readHexString(text, pos) {
  const end = text.indexOf('>', pos + 1)
  const stop = end === -1 ? text.length : end + 1
  return { value: { kind: 'token', text: text.slice(pos, stop) }, end: stop }
}

/** Read an array. */
function readArray(text, pos) {
  const items = []
  let i = pos + 1
  while (true) {
    i = skipWhitespace(text, i)
    if (i >= text.length) throw new Error('Unterminated PDF array')
    if (text[i] === ']') {
      i++
      break
    }
    const read = readValue(text, i)
    items.push(read.value)
    i = read.end
  }
  return { value: { kind: 'array', items }, end: i }
}

/** Read a dictionary, returning its entries keyed by full name (`'/Type'`). */
function readDictionary(text, pos) {
  const map = new Map()
  let i = pos + 2
  while (true) {
    i = skipWhitespace(text, i)
    if (i >= text.length) throw new Error('Unterminated PDF dictionary')
    if (text.startsWith('>>', i)) {
      i += 2
      break
    }
    if (text[i] !== '/') {
      throw new Error(`PDF dictionary key must be a name at byte ${i}`)
    }
    const key = readName(text, i)
    const read = readValue(text, key.end)
    map.set(key.value.text, read.value)
    i = read.end
  }
  return { value: { kind: 'dict', map }, end: i }
}

/** Read a number, a `N G R` reference, or one of the bare keywords. */
function readNumberOrKeyword(text, pos) {
  for (const keyword of ['true', 'false', 'null']) {
    if (text.startsWith(keyword, pos) && isDelimiter(text[pos + keyword.length])) {
      return { value: { kind: 'token', text: keyword }, end: pos + keyword.length }
    }
  }

  const match = /^[+-]?(?:\d+\.?\d*|\.\d+)/.exec(text.slice(pos, pos + 40))
  if (!match) {
    throw new Error(`Unsupported PDF token at byte ${pos}: ${JSON.stringify(text.slice(pos, pos + 24))}`)
  }
  const token = match[0]
  const end = pos + token.length

  if (/^\d+$/.test(token)) {
    const secondStart = skipWhitespace(text, end)
    const second = /^\d+/.exec(text.slice(secondStart, secondStart + 20))
    if (second) {
      const afterSecond = skipWhitespace(text, secondStart + second[0].length)
      if (text[afterSecond] === 'R' && isDelimiter(text[afterSecond + 1])) {
        return {
          value: { kind: 'ref', num: Number(token), gen: Number(second[0]) },
          end: afterSecond + 1
        }
      }
    }
  }

  return { value: { kind: 'token', text: token }, end }
}

/** Read any PDF object, returning its value and the position after it. */
function readValue(text, pos) {
  const i = skipWhitespace(text, pos)
  if (text.startsWith('<<', i)) return readDictionary(text, i)
  if (text[i] === '[') return readArray(text, i)
  if (text[i] === '/') return readName(text, i)
  if (text[i] === '(') return readLiteralString(text, i)
  if (text[i] === '<') return readHexString(text, i)
  return readNumberOrKeyword(text, i)
}

/** The `/Type` name of a value, whether it is a dictionary or a stream. */
function typeNameOf(value) {
  const map = value.kind === 'stream' ? value.dict.map : value.kind === 'dict' ? value.map : null
  if (!map) return null
  const type = map.get('/Type')
  return type && type.kind === 'token' ? type.text : null
}

/** The integer a token holds, or null when it is not a plain integer. */
function tokenInteger(value) {
  if (value && value.kind === 'token' && /^\d+$/.test(value.text)) return Number(value.text)
  return null
}

/**
 * Parse a PDF into its indirect objects.
 *
 * @param {Buffer} buffer
 * @returns {{objects: Map<number, object>, pages: Array<{num: number, value: object}>, version: string}}
 */
export function parsePdfDocument(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 8) {
    throw new Error('Invalid PDF: too short to be a PDF file')
  }
  const text = buffer.toString('latin1')
  const header = /%PDF-(\d+\.\d+)/.exec(text.slice(0, 1024))
  if (!header) throw new Error('Invalid PDF: the %PDF- header was not found')
  if (text.includes('/ObjStm')) {
    throw new Error('Unsupported PDF: compressed object streams (/ObjStm) are not handled')
  }

  // Candidate object headers. A candidate inside a stream is filtered out
  // below, because it falls inside a span that has already been parsed.
  const candidates = []
  for (const match of text.matchAll(/(\d{1,10})[ \t\r\n]+(\d{1,5})[ \t\r\n]+obj\b/g)) {
    const before = match.index === 0 ? ' ' : text[match.index - 1]
    if (!/[\s>\])]/.test(before)) continue
    candidates.push({ num: Number(match[1]), gen: Number(match[2]), offset: match.index })
  }

  const offsets = new Map(candidates.map((c) => [c.num, c.offset]))
  const objects = new Map()
  const resolving = new Set()

  /**
   * The byte length of a stream, from a direct `/Length` or through the object
   * it points at. A length object is a plain integer, so reading it on demand
   * cannot recurse into another stream.
   */
  const lengthOf = (dict) => {
    const value = dict.map.get('/Length')
    const direct = tokenInteger(value)
    if (direct !== null) return direct
    if (!value || value.kind !== 'ref') return null
    const cached = tokenInteger(objects.get(value.num))
    if (cached !== null) return cached
    const offset = offsets.get(value.num)
    if (offset === undefined || resolving.has(value.num)) return null
    resolving.add(value.num)
    try {
      const parsed = readIndirectObject(offset)
      return tokenInteger(parsed.value)
    } finally {
      resolving.delete(value.num)
    }
  }

  /** Parse the indirect object whose header starts at `offset`. */
  function readIndirectObject(offset) {
    const head = /^(\d+)[ \t\r\n]+(\d+)[ \t\r\n]+obj\b/.exec(text.slice(offset, offset + 40))
    if (!head) throw new Error(`Invalid PDF: no object header at byte ${offset}`)
    const body = readValue(text, offset + head[0].length)
    let end = body.end

    if (body.value.kind === 'dict') {
      const afterDict = skipWhitespace(text, body.end)
      if (text.startsWith('stream', afterDict) && isDelimiter(text[afterDict + 6])) {
        let dataStart = afterDict + 6
        if (text[dataStart] === '\r') dataStart++
        if (text[dataStart] === '\n') dataStart++

        const declared = lengthOf(body.value)
        let dataEnd = declared === null ? -1 : dataStart + declared
        // The declared length is authoritative only when `endstream` really is
        // there; a wrong or indirect-broken length must not eat the next object.
        if (dataEnd < 0 || !/^[\r\n\s]*endstream/.test(text.slice(dataEnd, dataEnd + 24))) {
          const marker = text.indexOf('endstream', dataStart)
          if (marker === -1) throw new Error(`Invalid PDF: a stream at byte ${offset} is never closed`)
          dataEnd = marker
          if (text[dataEnd - 2] === '\r' && text[dataEnd - 1] === '\n') dataEnd -= 2
          else if (text[dataEnd - 1] === '\n' || text[dataEnd - 1] === '\r') dataEnd -= 1
          // The length found this way is the true one, so the dict is corrected
          // rather than left claiming a length that would truncate the stream.
          body.value.map.set('/Length', { kind: 'token', text: String(Math.max(0, dataEnd - dataStart)) })
        }

        const closed = text.indexOf('endobj', dataEnd)
        return {
          num: Number(head[1]),
          gen: Number(head[2]),
          value: { kind: 'stream', dict: body.value, data: text.slice(dataStart, Math.max(dataStart, dataEnd)) },
          end: closed === -1 ? dataEnd : closed + 'endobj'.length
        }
      }
    }

    const closed = text.indexOf('endobj', end)
    if (closed !== -1) end = closed + 'endobj'.length
    return { num: Number(head[1]), gen: Number(head[2]), value: body.value, end }
  }

  let consumedUntil = 0
  for (const candidate of candidates) {
    // A header pattern inside a stream belongs to that stream's bytes.
    if (candidate.offset < consumedUntil) continue
    const parsed = readIndirectObject(candidate.offset)
    objects.set(parsed.num, parsed.value)
    consumedUntil = Math.max(consumedUntil, parsed.end)
  }

  // The catalog identifies the page tree; the type must be a name, not a
  // substring of one (`/CatalogX` is a different type). Every catalog and every
  // cross-reference stream is structural — a merge replaces them, so they are
  // not copied. A cross-reference stream is a stream object, so its `/Type` is
  // read from its dictionary.
  let catalog = null
  const structural = new Set()
  for (const [num, value] of objects) {
    const type = typeNameOf(value)
    if (type === '/Catalog') {
      structural.add(num)
      if (!catalog) catalog = value
    } else if (type === '/XRef') {
      structural.add(num)
    }
  }
  if (!catalog) throw new Error('Invalid PDF: no document catalog (/Type /Catalog) was found')

  const pages = collectPages(objects, catalog, structural)
  if (pages.length === 0) throw new Error('Invalid PDF: the page tree holds no pages')

  return { objects, pages, structural, version: header[1] }
}

/**
 * Walk the page tree from the catalog, in drawing order.
 *
 * While walking, the objects that make up the page tree itself and the
 * document catalog are recorded in `structural`: a merged document rebuilds
 * both, so copying the originals would leave a second, unreachable catalog
 * behind.
 *
 * @param {Map<number, object>} objects
 * @param {object} catalog
 * @param {Set<number>} structural - filled with catalog / page-tree / xref ids
 * @returns {Array<{num: number, value: object}>}
 */
function collectPages(objects, catalog, structural) {
  const pages = []
  const visited = new Set()

  const walk = (node, num) => {
    if (!node || node.kind !== 'dict') return
    const typeName = typeNameOf(node)
    if (typeName === '/Page') {
      pages.push({ num, value: node })
      return
    }
    // Anything else in the tree is an intermediate /Pages node.
    structural.add(num)
    const kids = node.map.get('/Kids')
    if (!kids || kids.kind !== 'array') return
    for (const kid of kids.items) {
      if (kid.kind !== 'ref' || visited.has(kid.num)) continue
      visited.add(kid.num)
      walk(objects.get(kid.num), kid.num)
    }
  }

  const root = catalog.map.get('/Pages')
  if (!root || root.kind !== 'ref') throw new Error('Invalid PDF: the catalog has no /Pages reference')
  walk(objects.get(root.num), root.num)
  return pages
}

/** Serialize a value, renumbering every reference through `renumber`. */
function serializeValue(value, renumber) {
  switch (value.kind) {
    case 'token':
      return value.text
    case 'ref': {
      // `absolute` marks a reference this module built *after* renumbering
      // (the merged page tree), so it is already in the target numbering.
      if (value.absolute === true) return `${value.num} ${value.gen} R`
      const mapped = renumber.get(value.num)
      if (mapped === undefined) {
        // Writing a dangling reference would produce a PDF that opens and then
        // fails on the page that needs it — fail where the cause is visible.
        throw new Error(`Invalid PDF: reference ${value.num} ${value.gen} R has no target object`)
      }
      return `${mapped} ${value.gen} R`
    }
    case 'array':
      return `[ ${value.items.map((item) => serializeValue(item, renumber)).join(' ')} ]`
    case 'dict':
      return serializeDictionary(value.map, renumber)
    case 'stream':
      // The dict is rewritten (references renumbered); the stream bytes are
      // copied exactly, so /Length stays the value it already claims.
      return `${serializeDictionary(value.dict.map, renumber)}\nstream\n${value.data}\nendstream`
    default:
      throw new Error(`Unsupported PDF value kind: ${value.kind}`)
  }
}

/** Serialize a dictionary's entries. */
function serializeDictionary(map, renumber) {
  const parts = []
  for (const [key, value] of map) parts.push(`${key} ${serializeValue(value, renumber)}`)
  return `<< ${parts.join(' ')} >>`
}

/**
 * The number of pages in a PDF.
 *
 * Parses the page tree; a PDF this module cannot parse is counted by looking
 * for `/Type /Page` markers instead of reporting zero.
 *
 * @param {Buffer} buffer
 * @returns {number}
 */
export function countPdfPages(buffer) {
  try {
    return parsePdfDocument(buffer).pages.length
  } catch {
    return (buffer.toString('latin1').match(/\/Type\s*\/Page(?![sA-Za-z])/g) || []).length
  }
}

/**
 * Concatenate PDFs into one document, in the given order.
 *
 * @param {Buffer[]} buffers
 * @returns {Buffer}
 */
export function mergePdfs(buffers) {
  if (!Array.isArray(buffers) || buffers.length === 0) {
    throw new Error('mergePdfs requires at least one PDF buffer')
  }
  // A single document needs no rewrite; returning it untouched keeps the
  // one-sheet case byte-identical to what the converter produced.
  if (buffers.length === 1) return Buffer.from(buffers[0])

  const documents = buffers.map((buffer) => parsePdfDocument(buffer))

  // Object numbers must be unique across the merged file, so each document's
  // non-structural objects get a contiguous block of the new numbering. The
  // documents' own catalogs and page trees are not copied: this module writes
  // one catalog and one page tree for the result.
  let next = 1
  const plans = documents.map((document) => {
    const renumber = new Map()
    for (const num of document.objects.keys()) {
      if (document.structural.has(num)) continue
      renumber.set(num, next++)
    }
    return { document, renumber }
  })

  const totalPages = plans.reduce((sum, plan) => sum + plan.document.pages.length, 0)
  const pagesNum = next++
  const catalogNum = next++

  // A page's /Parent must point at the merged page tree, not at the /Pages
  // node of the document it came from. The new page tree does not exist in any
  // input's numbering, so the reference is marked absolute.
  const kidRefs = []
  for (const plan of plans) {
    for (const page of plan.document.pages) {
      page.value.map.set('/Parent', { kind: 'ref', num: pagesNum, gen: 0, absolute: true })
      kidRefs.push(`${plan.renumber.get(page.num)} 0 R`)
    }
  }

  const chunks = []
  let offset = 0
  const offsets = new Map()
  const write = (text) => {
    chunks.push(text)
    offset += Buffer.byteLength(text, 'latin1')
  }

  write(`%PDF-${documents[0].version}\n%\u00e2\u00e3\u00cf\u00d3\n`)

  for (const plan of plans) {
    for (const [num, value] of plan.document.objects) {
      if (plan.document.structural.has(num)) continue
      const newNum = plan.renumber.get(num)
      offsets.set(newNum, offset)
      write(`${newNum} 0 obj\n${serializeValue(value, plan.renumber)}\nendobj\n`)
    }
  }

  offsets.set(pagesNum, offset)
  write(`${pagesNum} 0 obj\n<< /Type /Pages /Count ${totalPages} /Kids [ ${kidRefs.join(' ')} ] >>\nendobj\n`)
  offsets.set(catalogNum, offset)
  write(`${catalogNum} 0 obj\n<< /Type /Catalog /Pages ${pagesNum} 0 R >>\nendobj\n`)

  const xrefOffset = offset
  const entries = [`xref\n0 ${next}\n`, '0000000000 65535 f \n']
  for (let num = 1; num < next; num++) {
    const at = offsets.get(num)
    if (at === undefined) throw new Error(`Internal error: object ${num} was never written`)
    entries.push(`${String(at).padStart(10, '0')} 00000 n \n`)
  }
  entries.push(`trailer\n<< /Size ${next} /Root ${catalogNum} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`)
  write(entries.join(''))

  return Buffer.from(chunks.join(''), 'latin1')
}

/**
 * Point a workbook at one of its worksheets.
 *
 * `activeTab` is the tab `x2t` renders, so this is what selects the sheet for a
 * single-sheet conversion. The `<bookViews>` container is created when the
 * workbook has none.
 *
 * @param {string} workbookXml
 * @param {number} index - 0-based sheet index
 * @returns {string}
 */
export function setWorkbookActiveTab(workbookXml, index) {
  const tab = String(index)
  const views = extractElements(workbookXml, 'bookViews')[0]
  if (views) {
    const view = extractElements(views.outerXml, 'workbookView')[0]
    if (view) {
      // The attribute is replaced rather than appended, so a workbook that
      // already points at another tab cannot end up with two activeTab values.
      const rest = view.outerXml
        .replace(/^<workbookView/, '')
        .replace(/\/?>$/, '')
        .replace(/\s*activeTab\s*=\s*"[^"]*"/, '')
      const updated = `<workbookView activeTab="${tab}"${rest}/>`
      return workbookXml.replace(views.outerXml, views.outerXml.replace(view.outerXml, updated))
    }
    return workbookXml.replace(views.outerXml, `<bookViews><workbookView activeTab="${tab}"/></bookViews>`)
  }

  const bookViews = `<bookViews><workbookView activeTab="${tab}"/></bookViews>`
  // bookViews follows workbookPr (or fileVersion) and precedes sheets.
  const anchor = extractElements(workbookXml, 'workbookPr')[0]
    || extractElements(workbookXml, 'fileVersion')[0]
  if (anchor) {
    const at = anchor.index + anchor.outerXml.length
    return workbookXml.slice(0, at) + bookViews + workbookXml.slice(at)
  }
  const sheets = extractElements(workbookXml, 'sheets')[0]
  if (sheets) return workbookXml.slice(0, sheets.index) + bookViews + workbookXml.slice(sheets.index)
  return workbookXml.replace(/<workbook(?=[\s>])[^>]*>/, (open) => `${open}${bookViews}`)
}
