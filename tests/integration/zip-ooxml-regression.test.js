import { test, describe, before } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import zlib from 'node:zlib'
import { ZipArchive, crc32 } from '../../src/shared/zip.js'
import { DocxEngine } from '../../src/r7/docx.js'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { PptxEngine } from '../../src/r7/pptx.js'
import { R7Adapter } from '../../src/r7/adapter.js'

const tmpDir = path.join(os.tmpdir(), `dsh_r7_regression_${Date.now()}`)

/** True when every byte of two buffers matches. */
function sameBytes(a, b) {
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) return false
  return Buffer.compare(a, b) === 0
}

/**
 * Locate the first central directory header in an archive and return the
 * offset of a given field, so a test can deliberately corrupt one field.
 */
function centralFieldOffset(buffer, fieldOffset) {
  for (let i = 0; i + 46 <= buffer.length; i++) {
    if (buffer.readUInt32LE(i) === 0x02014b50) return i + fieldOffset
  }
  throw new Error('central directory header not found')
}

describe('ZIP / OOXML regression suite', () => {
  let adapter
  let r7Available = false
  let templates = {}

  before(async () => {
    fs.mkdirSync(tmpDir, { recursive: true })
    adapter = new R7Adapter()
    const info = await adapter.detect()
    r7Available = info.installed && Boolean(info.templatesPath)
    if (r7Available) {
      for (const type of ['docx', 'xlsx', 'pptx']) {
        const p = await adapter.getTemplatePath(type)
        if (p) templates[type] = p
      }
    }
  })

  describe('real R7 documents', () => {
    test('every detected native template opens as a valid archive', async (t) => {
      if (!r7Available || Object.keys(templates).length === 0) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      for (const [type, filePath] of Object.entries(templates)) {
        const zip = await ZipArchive.fromFile(filePath)
        assert.ok(zip.list().length > 0, `${type} template has members`)
        assert.ok(zip.has('[Content_Types].xml'), `${type} has [Content_Types].xml`)
      }
    })

    test('untouched documents survive a read/write cycle byte-for-byte', async (t) => {
      if (!r7Available || Object.keys(templates).length === 0) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      for (const [type, filePath] of Object.entries(templates)) {
        const original = fs.readFileSync(filePath)
        const zip = await ZipArchive.fromFile(filePath)
        const roundTripped = zip.toBuffer()

        // The strongest available guarantee: reading a package and writing it
        // back without edits must not alter a single byte, so no style,
        // theme, numbering or relationship part can drift.
        assert.ok(
          sameBytes(original, roundTripped),
          `${type} round-trip must be byte-identical (${original.length} vs ${roundTripped.length})`
        )
      }
    })

    test('editing one member leaves every other member byte-identical', async (t) => {
      if (!r7Available || !templates.docx) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const original = await ZipArchive.fromFile(templates.docx)
      const edited = await ZipArchive.fromFile(templates.docx)

      const documentXml = edited.getText('word/document.xml')
      edited.setText('word/document.xml', documentXml.replace('</w:body>', '<w:p><w:r><w:t>injected</w:t></w:r></w:p></w:body>'))

      // Re-parse the serialized result so every member carries the stream that
      // would actually be written to disk.
      const rewritten = ZipArchive.fromBuffer(edited.toBuffer())

      const before = original.list()
      const after = rewritten.list()
      assert.deepEqual(after, before, 'member list is preserved')

      for (const name of before) {
        const a = original.entries.get(name).raw
        const b = rewritten.entries.get(name).raw
        if (name === 'word/document.xml') {
          assert.ok(!sameBytes(a, b), 'the edited part must change')
        } else {
          assert.ok(sameBytes(a, b), `untouched member "${name}" must not change`)
        }
      }
    })

    test('document engines produce packages the archive can reopen', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const docxPath = path.join(tmpDir, 'roundtrip.docx')
      const docx = new DocxEngine(adapter)
      await docx.create(docxPath, {
        title: 'Round trip',
        paragraphs: ['Первый абзац', 'Второй абзац'],
        tables: [{ rows: [['A', 'B'], ['C', 'D']] }]
      })
      const validation = await docx.validate(docxPath)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))

      // Re-open and re-write: the result must still be a healthy document.
      const reopened = await ZipArchive.fromFile(docxPath)
      const rewritten = path.join(tmpDir, 'roundtrip-2.docx')
      await reopened.save(rewritten)
      const second = await docx.validate(rewritten)
      assert.equal(second.valid, true, JSON.stringify(second.errors))
    })
  })

  describe('Unicode handling', () => {
    test('Cyrillic, CJK and emoji survive a round-trip', () => {
      const zip = new ZipArchive()
      const content = 'Привет, мир! 你好世界 こんにちは مرحبا 🎉🚀 Ω≈ç√'
      zip.setText('docs/unicode.txt', content)

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.equal(reopened.getText('docs/unicode.txt'), content)
    })

    test('non-ASCII member names survive a round-trip', () => {
      const zip = new ZipArchive()
      const names = [
        'документы/отчёт.txt',
        '文档/说明.txt',
        'emoji/🎉.txt'
      ]
      for (const n of names) zip.setText(n, `payload for ${n}`)

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      for (const n of names) {
        assert.ok(reopened.has(n), `member "${n}" must survive`)
        assert.equal(reopened.getText(n), `payload for ${n}`)
      }
    })

    test('UTF-8 content is byte-exact, not re-encoded', () => {
      const zip = new ZipArchive()
      const text = 'Тест Юникода: ЁЖЗИЙ'
      zip.setText('u.txt', text)
      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.ok(sameBytes(Buffer.from(text, 'utf8'), Buffer.from(reopened.getText('u.txt'), 'utf8')))
    })

    test('Unicode reaches the DOCX/OOXML text layer', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const target = path.join(tmpDir, 'unicode.docx')
      const docx = new DocxEngine(adapter)
      await docx.create(target, {
        title: 'Отчёт 2026 — 中文 — 🎯',
        paragraphs: ['Кириллица', '汉字测试', 'emoji 🎉']
      })
      const read = await docx.read(target, { format: 'markdown' })
      assert.ok(read.content.includes('Отчёт 2026'), 'Cyrillic title')
      assert.ok(read.content.includes('中文') || read.content.includes('汉字测试'), 'CJK paragraph')
      assert.ok(read.content.includes('🎉'), 'emoji paragraph')
    })
  })

  describe('compression methods', () => {
    test('stored entries round-trip without deflating', () => {
      const zip = new ZipArchive()
      const payload = Buffer.from('stored payload '.repeat(20), 'utf8')
      zip.setStored('stored.bin', payload)
      zip.setText('deflated.txt', 'deflated payload '.repeat(20))

      const buffer = zip.toBuffer()
      const reopened = ZipArchive.fromBuffer(buffer)

      assert.ok(sameBytes(reopened.getBuffer('stored.bin'), payload))
      assert.equal(reopened.getMethod('stored.bin'), 0, 'stored entry keeps method 0')
      assert.equal(reopened.getMethod('deflated.txt'), 8, 'compressible entry uses deflate')
    })

    test('both methods coexist in one archive', () => {
      const zip = new ZipArchive()
      zip.setStored('a.bin', Buffer.from('a'.repeat(500)))
      zip.setText('b.txt', 'b'.repeat(500))
      zip.setStored('c.bin', Buffer.alloc(0))

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.equal(reopened.list().length, 3)
      assert.equal(reopened.getMethod('a.bin'), 0)
      assert.equal(reopened.getMethod('b.txt'), 8)
      assert.equal(reopened.getBuffer('c.bin').length, 0)
    })

    test('incompressible data falls back to stored instead of growing', () => {
      const zip = new ZipArchive()
      const random = Buffer.alloc(4096)
      for (let i = 0; i < random.length; i++) random[i] = (i * 37 + 11) % 251
      zip.setBuffer('random.bin', random)

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.ok(sameBytes(reopened.getBuffer('random.bin'), random))
      assert.ok(
        reopened.toBuffer().length <= random.length + 400,
        'archive must not be dramatically larger than its payload'
      )
    })

    test('deflate stream produced by the writer is valid raw deflate', () => {
      const zip = new ZipArchive()
      const text = 'compress me '.repeat(100)
      zip.setText('c.txt', text)
      const raw = zip.entries.get('c.txt')
      assert.equal(raw.raw, null, 'a newly written entry has no replayed stream yet')

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      const mixed = reopened.entries.get('c.txt')
      assert.ok(sameBytes(zlib.inflateRawSync(mixed.raw), Buffer.from(text, 'utf8')))
    })
  })

  describe('large documents', () => {
    test('an archive with many members round-trips', () => {
      const zip = new ZipArchive()
      const count = 2000
      for (let i = 0; i < count; i++) {
        zip.setText(`part/${i}.txt`, `payload number ${i}`)
      }
      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.equal(reopened.list().length, count)
      assert.equal(reopened.getText('part/0.txt'), 'payload number 0')
      assert.equal(reopened.getText(`part/${count - 1}.txt`), `payload number ${count - 1}`)
    })

    test('a multi-megabyte member round-trips intact', () => {
      const zip = new ZipArchive()
      const big = Buffer.alloc(4 * 1024 * 1024)
      for (let i = 0; i < big.length; i++) big[i] = i % 256
      zip.setBuffer('big.bin', big)

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      const restored = reopened.getBuffer('big.bin')
      assert.equal(restored.length, big.length)
      assert.ok(sameBytes(restored, big))
    })

    test('a large DOCX with many paragraphs stays valid', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const target = path.join(tmpDir, 'large.docx')
      const paragraphs = []
      for (let i = 0; i < 500; i++) {
        paragraphs.push(`Раздел ${i}. Строка содержимого для нагрузочной проверки документа.`)
      }
      const docx = new DocxEngine(adapter)
      await docx.create(target, { title: 'Большой документ', paragraphs })

      const validation = await docx.validate(target)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))

      const inspection = await docx.inspect(target)
      assert.ok(inspection.paragraphsCount >= 500, `expected >= 500 paragraphs, got ${inspection.paragraphsCount}`)
    })

    test('large archive CRC-32 stays consistent', () => {
      const payload = Buffer.alloc(1024 * 1024, 7)
      const zip = new ZipArchive()
      zip.setBuffer('crc.bin', payload)
      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.equal(reopened.entries.get('crc.bin').crc, crc32(payload))
    })
  })

  describe('malformed packages', () => {
    test('an empty buffer is rejected with a clear error', () => {
      assert.throws(() => ZipArchive.fromBuffer(Buffer.alloc(0)), /Invalid ZIP/)
    })

    test('a buffer shorter than the EOCD record is rejected', () => {
      assert.throws(() => ZipArchive.fromBuffer(Buffer.from('not a zip at all')), /Invalid ZIP/)
    })

    test('data without an EOCD record is rejected', () => {
      const junk = Buffer.alloc(64, 0x41)
      assert.throws(() => ZipArchive.fromBuffer(junk), /end-of-central-directory record not found/)
    })

    test('a truncated archive (central directory cut off) is rejected', () => {
      const zip = new ZipArchive()
      zip.setText('a.txt', 'hello')
      zip.setText('b.txt', 'world')
      const full = zip.toBuffer()
      const truncated = full.subarray(0, full.length - 30)
      assert.throws(() => ZipArchive.fromBuffer(truncated), /Invalid ZIP/)
    })

    test('a corrupted local header signature is detected', () => {
      const zip = new ZipArchive()
      zip.setText('a.txt', 'hello')
      const buffer = zip.toBuffer()
      // Corrupt the first local file header signature.
      buffer.writeUInt32LE(0xdeadbeef, 0)
      assert.throws(() => ZipArchive.fromBuffer(buffer), /bad local header signature/)
    })

    test('a corrupted compressed payload is detected', () => {
      const zip = new ZipArchive()
      zip.setText('a.txt', 'x'.repeat(2000))
      const buffer = zip.toBuffer()
      // Flip bytes inside the deflate stream (past the 30-byte local header).
      for (let i = 40; i < 60; i++) buffer[i] = 0xff
      assert.throws(() => ZipArchive.fromBuffer(buffer), /Invalid ZIP/)
    })

    test('an unsupported compression method is rejected', () => {
      const zip = new ZipArchive()
      zip.setText('a.txt', 'hello')
      const buffer = zip.toBuffer()
      // The method field lives at +10 of both the local and the central
      // header; patch both so the mismatch check does not fire first.
      buffer.writeUInt16LE(99, 8)
      buffer.writeUInt16LE(99, centralFieldOffset(buffer, 10))
      assert.throws(() => ZipArchive.fromBuffer(buffer), /unsupported compression method/)
    })

    test('a local/central compression method mismatch is detected', () => {
      const zip = new ZipArchive()
      // Highly compressible payload, so the writer definitely chooses deflate.
      zip.setText('a.txt', 'x'.repeat(2000))
      const buffer = zip.toBuffer()
      assert.equal(buffer.readUInt16LE(8), 8, 'precondition: local header says deflate')

      // Patch only the central directory to claim "stored".
      buffer.writeUInt16LE(0, centralFieldOffset(buffer, 10))
      assert.throws(() => ZipArchive.fromBuffer(buffer), /compression method mismatch/)
    })

    test('a Zip64 archive is rejected with an explicit message', () => {
      const eocd = Buffer.alloc(22)
      eocd.writeUInt32LE(0x06054b50, 0)
      eocd.writeUInt16LE(0xffff, 10) // entry count sentinel
      eocd.writeUInt32LE(0xffffffff, 16)
      const locator = Buffer.alloc(20)
      locator.writeUInt32LE(0x07064b50, 0)
      assert.throws(() => ZipArchive.fromBuffer(Buffer.concat([locator, eocd])), /Zip64 archives are not supported/)
    })

    test('a non-Buffer input is rejected', () => {
      assert.throws(() => ZipArchive.fromBuffer('a string'), /expected a Buffer/)
    })

    test('a document engine reports a corrupt DOCX rather than throwing', async () => {
      const corrupt = path.join(tmpDir, 'corrupt.docx')
      fs.writeFileSync(corrupt, Buffer.alloc(120, 0x00))
      const docx = new DocxEngine(adapter)
      const result = await docx.validate(corrupt)
      assert.equal(result.valid, false)
      assert.ok(result.errors.length > 0)
    })

    test('inspect on a corrupt file throws a descriptive error', async () => {
      const corrupt = path.join(tmpDir, 'corrupt2.docx')
      fs.writeFileSync(corrupt, Buffer.from('definitely not a docx'))
      const docx = new DocxEngine(adapter)
      await assert.rejects(() => docx.inspect(corrupt), /Invalid ZIP/)
    })
  })

  describe('round-trip fidelity', () => {
    test('reading and writing a synthesized package is idempotent', () => {
      const zip = new ZipArchive()
      zip.setText('one.txt', 'first')
      zip.setText('dir/two.txt', 'second')
      zip.setStored('three.bin', Buffer.from([1, 2, 3, 4, 5]))

      const once = zip.toBuffer()
      const twice = ZipArchive.fromBuffer(once).toBuffer()
      const thrice = ZipArchive.fromBuffer(twice).toBuffer()

      assert.ok(sameBytes(once, twice), 'second write must equal the first')
      assert.ok(sameBytes(twice, thrice), 'third write must equal the second')
    })

    test('an edited member changes while its siblings stay frozen', () => {
      const zip = new ZipArchive()
      zip.setText('keep-a.txt', 'alpha')
      zip.setText('edit-me.txt', 'beta')
      zip.setStored('keep-b.bin', Buffer.from('gamma'))

      const original = zip.toBuffer()
      const reopened = ZipArchive.fromBuffer(original)
      reopened.setText('edit-me.txt', 'beta-changed')
      const after = reopened.toBuffer()

      const before = ZipArchive.fromBuffer(original)
      const updated = ZipArchive.fromBuffer(after)

      assert.ok(sameBytes(before.entries.get('keep-a.txt').raw, updated.entries.get('keep-a.txt').raw))
      assert.ok(sameBytes(before.entries.get('keep-b.bin').raw, updated.entries.get('keep-b.bin').raw))
      assert.ok(!sameBytes(before.entries.get('edit-me.txt').raw, updated.entries.get('edit-me.txt').raw))
      assert.equal(updated.getText('edit-me.txt'), 'beta-changed')
    })

    test('member order is preserved across a round-trip', () => {
      const zip = new ZipArchive()
      const names = ['z.txt', 'a.txt', 'm/n.txt', 'b.txt']
      for (const n of names) zip.setText(n, n)
      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.deepEqual(reopened.list(), names)
    })

    test('removing a member leaves the rest intact', () => {
      const zip = new ZipArchive()
      zip.setText('a.txt', 'aaa')
      zip.setText('b.txt', 'bbb')
      zip.setText('c.txt', 'ccc')

      assert.equal(zip.remove('b.txt'), true)
      assert.equal(zip.remove('b.txt'), false, 'removing twice reports false')

      const reopened = ZipArchive.fromBuffer(zip.toBuffer())
      assert.deepEqual(reopened.list(), ['a.txt', 'c.txt'])
      assert.equal(reopened.getText('a.txt'), 'aaa')
      assert.equal(reopened.getText('c.txt'), 'ccc')
    })

    test('OOXML packages survive an independent rewrite of every member', async (t) => {
      if (!r7Available || !templates.xlsx) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const original = await ZipArchive.fromFile(templates.xlsx)
      const rewritten = new ZipArchive()

      // Copy every member through the setter path, forcing re-compression and
      // recomputed CRCs instead of a verbatim replay.
      for (const name of original.list()) {
        rewritten.setBuffer(name, original.getBuffer(name))
      }

      const target = path.join(tmpDir, 'rewritten.xlsx')
      await rewritten.save(target)

      const xlsx = new XlsxEngine(adapter)
      const validation = await xlsx.validate(target)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))

      // Content must be identical even though the container was rebuilt.
      const reopened = await ZipArchive.fromFile(target)
      for (const name of original.list()) {
        assert.ok(
          sameBytes(original.getBuffer(name), reopened.getBuffer(name)),
          `member "${name}" content must survive a full rewrite`
        )
      }
    })
  })

  describe('format-specific regression', () => {
    test('XLSX engine keeps untouched sheets byte-identical', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const source = path.join(tmpDir, 'sheets-source.xlsx')
      const xlsx = new XlsxEngine(adapter)
      await xlsx.create(source, { sheets: [{ name: 'Лист1', data: [['A', 'B'], [1, 2]] }] })

      const before = await ZipArchive.fromFile(source)
      await xlsx.write(source, { cells: [{ ref: 'C1', value: 'new' }] })
      const after = await ZipArchive.fromFile(source)

      // Shared strings and workbook metadata must not be rewritten by a cell write.
      for (const name of before.list()) {
        if (name === 'xl/worksheets/sheet1.xml') continue
        assert.ok(
          sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
          `"${name}" must be untouched by a single-cell write`
        )
      }
    })

    test('PPTX engine preserves non-slide parts on an edit', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const source = path.join(tmpDir, 'slides-source.pptx')
      const pptx = new PptxEngine(adapter)
      await pptx.create(source, { title: 'Исходный заголовок' })

      const before = await ZipArchive.fromFile(source)
      await pptx.editSlide(source, { title: 'Новый заголовок' })
      const after = await ZipArchive.fromFile(source)

      for (const name of before.list()) {
        if (name === 'ppt/slides/slide1.xml') continue
        assert.ok(
          sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
          `"${name}" must be untouched by a slide text edit`
        )
      }
      assert.equal(after.getText('ppt/slides/slide1.xml').includes('Новый заголовок'), true)
    })

    test('DOCX replace leaves unrelated parts frozen', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const source = path.join(tmpDir, 'replace-source.docx')
      const docx = new DocxEngine(adapter)
      await docx.create(source, { paragraphs: ['Исходный текст для замены.'] })

      const before = await ZipArchive.fromFile(source)
      await docx.replaceText(source, 'Исходный текст для замены.', 'Заменённый текст.')
      const after = await ZipArchive.fromFile(source)

      for (const name of before.list()) {
        if (name === 'word/document.xml') continue
        assert.ok(
          sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
          `"${name}" must be untouched by a text replace`
        )
      }
    })
  })
})
