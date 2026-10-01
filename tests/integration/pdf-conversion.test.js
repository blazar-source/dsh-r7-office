import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'

import { R7Adapter } from '../../src/r7/adapter.js'
import { inspectPdf } from '../../src/r7/pdf-inspect.js'

// ---------------------------------------------------------------------------
// Minimal, self-contained DOCX writer (no project code) so this test proves the
// CONVERTER works rather than proving the generator agrees with itself.
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

function buildDocx(paragraphs) {
  const body = paragraphs
    .map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)
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

const RUSSIAN = 'Привет, мир! Это проверка кириллицы.'

// ---------------------------------------------------------------------------

test('r7_convert produces a PDF with real extractable Cyrillic text', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!info.x2tPath) return t.skip('R7 x2t converter is not installed')
  if (!info.allFontsPath) {
    return t.skip('no R7 AllFonts.js on this host (R7-Office Desktop has never run)')
  }

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'r7-pdf-integration-'))
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'cyrillic.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN, 'Вторая строка текста.']))

  const result = await adapter.convert(source, target)

  assert.equal(result.success, true)
  assert.equal(result.mode, 'params-xml', 'the font-list aware params-XML form must be used')
  assert.equal(result.allFontsPath, info.allFontsPath)

  const quality = result.pdfTextQuality
  assert.ok(quality.textGlyphs > 0, 'the PDF must draw text, not empty glyph fills')
  assert.equal(quality.emptyFillOperators, 0, 'no glyph fill may paint nothing')
  assert.equal(quality.textLossSuspected, false)
  assert.notEqual(quality.verdict, 'outlined')
  assert.ok(quality.toUnicodeMappings > 0, 'body text must carry a ToUnicode map')
  assert.ok(quality.cyrillicMappings > 0, 'the ToUnicode map must cover Cyrillic')
  assert.deepEqual(quality.malformedToUnicode, [], 'ToUnicode counts must be well formed after repair')
})

test('r7_convert fails loudly instead of returning a text-less PDF', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!info.x2tPath) return t.skip('R7 x2t converter is not installed')

  // R7's installer ships a 0-byte AllFonts.js stub; forcing it reproduces the
  // original defect on any host that has the installation.
  const stub = info.installPath
    ? path.join(info.installPath, 'editors', 'sdkjs', 'common', 'AllFonts.js')
    : null
  if (!stub || !fs.existsSync(stub) || fs.statSync(stub).size >= 1024) {
    return t.skip('no 0-byte AllFonts.js stub in this installation')
  }

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'r7-pdf-integration-'))
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'broken.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN]))

  await assert.rejects(
    () => adapter.convert(source, target, { allFontsPath: stub }),
    (error) => {
      assert.match(error.message, /no extractable text/)
      assert.match(error.message, /AllFonts\.js/)
      assert.ok(error.pdfTextQuality, 'the quality report must travel with the error')
      assert.equal(error.pdfTextQuality.textLossSuspected, true)
      return true
    }
  )
})

test('r7_convert keeps the unusable PDF only when explicitly allowed', async (t) => {
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!info.x2tPath) return t.skip('R7 x2t converter is not installed')

  const stub = info.installPath
    ? path.join(info.installPath, 'editors', 'sdkjs', 'common', 'AllFonts.js')
    : null
  if (!stub || !fs.existsSync(stub) || fs.statSync(stub).size >= 1024) {
    return t.skip('no 0-byte AllFonts.js stub in this installation')
  }

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'r7-pdf-integration-'))
  const source = path.join(dir, 'cyrillic.docx')
  const target = path.join(dir, 'outlined.pdf')
  fs.writeFileSync(source, buildDocx([RUSSIAN]))

  const result = await adapter.convert(source, target, { allFontsPath: stub, allowOutlinedPdf: true })
  assert.equal(result.success, true)
  assert.equal(result.pdfTextQuality.textLossSuspected, true)
  assert.equal(inspectPdf(target).verdict !== 'text', true)
})
