import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'

import { R7Adapter } from '../../src/r7/adapter.js'
import { inspectPdf } from '../../src/r7/pdf-inspect.js'
import { parseFontListPaths } from '../../src/r7/font-substitutes.js'

// ---------------------------------------------------------------------------
// Minimal, self-contained DOCX writer (no project code) so this test proves the
// CONVERTER is right rather than proving the generator agrees with itself.
// ---------------------------------------------------------------------------

function crc32(buf) {
  let crc = 0xffffffff
  for (let i = 0; i < buf.length; i++) {
    let c = (crc ^ buf[i]) & 0xff
    for (let k = 0; k < 8; k++) c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1
    crc = (crc >>> 8) ^ c
  }
  return (crc ^ 0xffffffff) >>> 0
}

function zip(files) {
  const parts = []
  const central = []
  let offset = 0
  for (const [name, content] of files) {
    const data = Buffer.from(content, 'utf8')
    const compressed = zlib.deflateRawSync(data)
    const payload = compressed.length < data.length ? compressed : data
    const method = compressed.length < data.length ? 8 : 0
    const nameBuf = Buffer.from(name, 'utf8')
    const crc = crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(method, 8)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(payload.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBuf.length, 26)
    parts.push(local, nameBuf, payload)
    const entry = Buffer.alloc(46)
    entry.writeUInt32LE(0x02014b50, 0)
    entry.writeUInt16LE(20, 4)
    entry.writeUInt16LE(20, 6)
    entry.writeUInt16LE(method, 10)
    entry.writeUInt32LE(crc, 16)
    entry.writeUInt32LE(payload.length, 20)
    entry.writeUInt32LE(data.length, 24)
    entry.writeUInt16LE(nameBuf.length, 28)
    entry.writeUInt32LE(offset, 42)
    central.push(entry, nameBuf)
    offset += local.length + nameBuf.length + payload.length
  }
  const centralBuf = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(files.length, 8)
  end.writeUInt16LE(files.length, 10)
  end.writeUInt32LE(centralBuf.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...parts, centralBuf, end])
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

/**
 * Build a DOCX whose runs ask for a specific font.
 *
 * The font is the whole point of this test. The acceptance document asks (via
 * its theme) for "Liberation Sans", the metric-compatible stand-in R7 ships for
 * Arial: x2t then numbers the glyphs with Liberation Sans and embeds Arial, so
 * requesting any other family would make the test vacuous.
 */
function buildDocx(paragraphs, font = 'Liberation Sans') {
  const rPr = font
    ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/></w:rPr>`
    : ''
  const body = paragraphs
    .map((text) => `<w:p><w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)
    .join('')
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  return zip([
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', rels],
    ['word/document.xml', document]
  ])
}

/**
 * Text chosen so that the defect is unmissable: Arial and Liberation Sans number
 * every one of these characters differently, and they exercise the irregular
 * offsets too (« » · — and ё are not a constant shift away from their Arial glyphs).
 */
const RUSSIAN = 'ООО «Ромашка» — отчёт о внедрении: ЁЖ, ёж, №1, 55% и 100% готовности.'

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'docx-pdf-glyphs-'))
}

const r7Available = (await new R7Adapter().detect()).installed

test('the converted DOCX names the characters its glyphs actually draw', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!r7Available) return t.skip('R7 not installed')
  if (!info.allFontsPath) return t.skip('no R7 AllFonts.js on this host')

  const dir = tmpDir()
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'cyrillic.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN, 'Вторая строка текста для верности.']))

  const result = await adapter.convert(source, target)
  const quality = result.pdfTextQuality

  // The trap: all of this already held while the page rendered garbage.
  assert.ok(quality.cyrillicMappings > 0)
  assert.equal(quality.verdict, 'text')

  // ...and this is the assertion that actually catches it.
  assert.equal(quality.glyphChainChecked, true, 'the chain must be verifiable for a produced PDF')
  assert.equal(quality.glyphChainMismatches, 0, JSON.stringify(quality.glyphChain.fonts.filter((f) => f.mismatches)))
  assert.equal(quality.glyphMapInconsistent, false)
  assert.ok(quality.glyphChain.cidsChecked >= 20, 'the check must cover the body text, not a handful of glyphs')
  assert.ok(quality.glyphChain.sourceFontsMatched >= 1, 'at least one embedded font must have been identified')

  // And independently of the adapter's own report.
  const report = inspectPdf(target)
  assert.equal(report.glyphMapInconsistent, false)
  assert.equal(report.textGlyphs, quality.textGlyphs)
})

test('the font list handed to x2t redirects substitutes at the real fonts', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!r7Available) return t.skip('R7 not installed')
  if (!info.allFontsPath) return t.skip('no R7 AllFonts.js on this host')

  const plan = adapter.prepareFontList(info.allFontsPath)
  assert.ok(plan, 'the font list must be readable')
  assert.ok(plan.fontPaths.length > 0)
  if (!plan.sanitized) {
    return t.skip('this host has no metric-compatible substitute whose real font is installed')
  }

  const rewritten = fs.readFileSync(plan.path, 'utf8')
  assert.equal(parseFontListPaths(rewritten).length, plan.fontPaths.length, 'the list keeps its shape')
  for (const rewrite of plan.rewrites) {
    assert.notEqual(rewrite.from, rewrite.to)
    assert.equal(rewritten.includes(`"${rewrite.from}"`), false, 'the shadowing path must be gone')
    assert.equal(fs.existsSync(rewrite.to), true, 'every rewrite must point at a font that exists')
  }
  assert.match(rewritten, /__all_fonts_js_version__/, 'the file must stay a valid font list')

  const dir = tmpDir()
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'cyrillic.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN]))
  const result = await adapter.convert(source, target)
  assert.equal(result.mode, 'params-xml')
  assert.equal(result.fontList.used, plan.path)
  assert.equal(result.allFontsPath, info.allFontsPath, 'allFontsPath still reports the discovered list')
})

test('an unrewritten font list is detected and refused, not shipped', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!r7Available) return t.skip('R7 not installed')
  if (!info.allFontsPath) return t.skip('no R7 AllFonts.js on this host')
  if (!adapter.prepareFontList(info.allFontsPath).sanitized) {
    return t.skip('this host has no substitute shadowing to reproduce the defect with')
  }

  const dir = tmpDir()
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'broken.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN]))

  const options = { sanitizeFontList: false }
  let refused = null
  try {
    const result = await adapter.convert(source, target, options)
    // Not every host/library combination still produces the defect; if it did not,
    // there is nothing to refuse and the assertion below would be vacuous.
    assert.equal(result.pdfTextQuality.glyphMapInconsistent, false)
    return t.skip('this x2t build did not reproduce the substitute defect')
  } catch (error) {
    refused = error
  }

  assert.match(refused.message, /wrong glyphs/)
  assert.match(refused.message, /allowGlyphMismatchPdf/)
  assert.ok(refused.pdfTextQuality, 'the quality report must travel with the error')
  assert.equal(refused.pdfTextQuality.glyphMapInconsistent, true)
  assert.ok(refused.pdfTextQuality.glyphChainMismatches > 0)
  assert.equal(refused.pdfTextQuality.verdict, 'text', 'the text layer is fine - that is the whole trap')

  // Opting in keeps the PDF and reports exactly what is wrong with it.
  const kept = await adapter.convert(source, target, { ...options, allowGlyphMismatchPdf: true })
  assert.equal(kept.success, true)
  assert.equal(kept.pdfTextQuality.glyphMapInconsistent, true)
  assert.ok(kept.pdfTextQuality.glyphChainMismatches > 0)
})
